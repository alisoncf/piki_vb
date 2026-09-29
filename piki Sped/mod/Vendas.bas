Attribute VB_Name = "Vendas"
Option Explicit
Dim strSQL As String
Public bUsaVendaSupermercado As Boolean
Public UltimoVeiculoProcurado As Veiculo
Type ClienteVenda
    cod_cli As String
    cpfcnpj As String
    nomecli_ven As String
End Type
Type tipo_venda
    codigo As String
    descricao As String
End Type
Enum tipo_operacao_venda
    venda = 1
    Devolucao = 2
    Remessa = 3
    retorno_remessa = 4
End Enum



Sub verifica_atacado(num_ven As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rAtacado As New ADODB.Recordset
    strSQL = "UPDATE itvendas SET unit_ite="
    strSQL = strSQL & " (SELECT pv_pro FROM produtos WHERE cod_pro=itvendas.cod_pro)"
    strSQL = strSQL & " Where num_ven = '" & num_ven & "'"
    conn.Execute strSQL
    
    strSQL = "UPDATE itvendas SET final_ite=(qtd_ite*unit_ite)-desconto_ite,atacado_ite='N'"
    strSQL = strSQL & " Where num_ven = '" & num_ven & "'"
    conn.Execute strSQL
    
    
    
    
    strSQL = "select tc.cod_cat,tc.nome_cat,tc.descmax_cat,sum(ti.qtd_ite) as qtd, tc.qtdmin_cat "
    strSQL = strSQL & " from " & TAB_ITVENDAS & " as ti "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " inner join " & TAB_CATEGORIAS & " as tc on tc.cod_cat=tp.cod_cat "
    
    strSQL = strSQL & " where ti.num_ven='" & num_ven & "' "
    strSQL = strSQL & " and qtdmin_cat>0"
    strSQL = strSQL & " group by tc.cod_cat, tc.nome_cat "
    strSQL = strSQL & " HAVING qtd>=qtdmin_cat "
    
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        
        Do While Not .EOF
'            strSQL = "update " & TAB_ITVENDAS & " as itvendas set unit_ite = "
'            strSQL = strSQL & " (SELECT pv_pro-(pv_pro*" & F_GravaDec(!descmax_cat) & "/100)"
'            strSQL = strSQL & " FROM produtos WHERE cod_pro=itvendas.cod_pro )"
'            strSQL = strSQL & " where num_ven='" & num_ven & "'"
'            strSQL = strSQL & " and id in (select ti.id from " & TAB_ITVENDAS & " as ti inner join " & TAB_PRODUTOS & " as tp "
'            strSQL = strSQL & " on tp.cod_pro=ti.cod_pro where ti.num_ven='" & num_ven & "' and tp.cod_cat='" & !cod_cat & "')"
'
            strSQL = "select ti.id,tp.cod_pro,tp.pv_pro,(pv_pro-(pv_pro*descmax_cat/100)) as pa_pro,tc.cod_cat,tc.nome_cat,tc.descmax_cat,tc.qtdmin_cat "
            strSQL = strSQL & " from " & TAB_ITVENDAS & " as ti "
            strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp "
            strSQL = strSQL & " on tp.cod_pro=ti.cod_pro "
            strSQL = strSQL & " inner join " & TAB_CATEGORIAS & " as tc on tc.cod_cat=tp.cod_cat "
    
            strSQL = strSQL & " where ti.num_ven='" & num_ven & "' "
            strSQL = strSQL & " and tc.cod_cat='" & !cod_cat & "' "
            
            abreObjetoAdo rAtacado, strSQL
            With rAtacado
                Do While Not .EOF
                    strSQL = "update " & TAB_ITVENDAS & " set unit_ite=" & F_GravaDec(!pa_pro) & ","
                    strSQL = strSQL & " atacado_ite='S'"
                    strSQL = strSQL & " where id=" & !id
                    conn.Execute strSQL
                    .MoveNext
                Loop
                .Close
            End With
            
            
            
            
            .MoveNext
        Loop
        
        .Close
    
    End With
    
    strSQL = "UPDATE itvendas SET desconto_ite=(desconto_aplicado/100*(qtd_ite*unit_ite)), final_ite=((qtd_ite*unit_ite)-(desconto_aplicado/100*(qtd_ite*unit_ite)))"
    strSQL = strSQL & " Where num_ven = '" & num_ven & "'"
    conn.Execute strSQL
    
    
End Sub

Sub DeletaVenFec(num_ven As String)
    Dim strSQL As String
    strSQL = "delete from " & TAB_VENFEC
    strSQL = strSQL & " where num_ven='" & num_ven & "'"
    conn.Execute strSQL
     
End Sub
Function VerificaEmpresasDiferentesVenda(num_ven As String) As Integer
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select distinct count(cod_emp) as q from " & TAB_ITVENDAS & " where num_ven='" & num_ven & "'"
    abreObjetoAdo rPesq, strSQL
    
    
    VerificaEmpresasDiferentesVenda = rPesq!q
    rPesq.Close
trata_Erro:

End Function
Function CancelarVenda(num_ven As String) As Integer

    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    If vendas.F_VerificaVendaAberta(num_ven) <> "" Then
        CancelarVenda = 9
        Exit Function
    End If
    strSQL = "select tv.num_ven,situacao_ven, ti.* from " & TAB_VENDAS & " as tv left join " & TAB_ITVENDAS & " as ti "
    strSQL = strSQL & " on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " where tv.num_ven='" & num_ven & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            CancelarVenda = 1 'não encontrada
            .Close
            Exit Function
        End If
        
        
        
        If !situacao_ven > 1 Then
            CancelarVenda = 2 'VENDA NÃO ESTÁ ABERTA
            .Close
            Exit Function
        End If
        Do While Not .EOF
            If Not IsNull(!seq_ite) Then
                strSQL = "delete from " & TAB_ITVENDAS & " where seq_ite=" & !seq_ite & " "
            
                strSQL = strSQL & " and num_ven='" & num_ven & "' "
                conn.Execute strSQL
                
            
                
                If !situacao_ven <= 2 Then
                
                    strSQL = "update " & TAB_PRODUTOS & " set reservado_pro=reservado_pro - " & F_GravaDec(CStr(!qtd_ite))
                    strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
                    
                    conn.Execute strSQL
                    strSQL = "update estoque set reservado=reservado - " & F_GravaDec(CStr(!qtd_ite))
                    strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
                    strSQL = strSQL & " and cod_emp='" & !cod_emp & "'"
                    conn.Execute strSQL
                End If
                
            End If
        
            .MoveNext
        Loop
        
    End With
    strSQL = "delete from " & TAB_VENDAS & " where num_ven='" & num_ven & "' "
    conn.Execute strSQL
    
    CancelarVenda = 0 'sucesso
    Exit Function
trata_Erro:
    CancelarVenda = 1 'falha geral

End Function

Function F_GeraNovaVenda(Optional cod_cli As String = "00000", Optional cod_For As String = "06", _
Optional cpf As String = "", Optional tipo_ven As String = "V", Optional sCodEmp As String = "01") As String
    Dim sNum As String
    Dim sIp As String
    On Error GoTo trata_Erro
    'estoque cod_For='06
    'Sleep 100
    If strCod_Tur = "" Then
        strCod_Tur = "1"
    End If
    
    
    
    UltimoVeiculoProcurado.id = 0
    UltimoVeiculoProcurado.descricao = ""
    
    sNum = ChaveSQL(TAB_VENDAS, "num_ven", "000000#")
    strSQL = F_BuscaUmCampo("num_ven", "tab_Vendas", "num_ven='" & sNum & "'")
    If strSQL <> "" Then
        F_GeraNovaVenda = ""
        MsgBox "Já existe uma venda com o número " & strSQL, vbExclamation
        Exit Function
        
    End If
    
    strSQL = "insert into " & TAB_VENDAS & " (num_ven,cod_cai,cod_tur,cod_cli,ip_ven,data_ven,abrevenda_usu,cod_for,cpfcli_ven,tipo_ven,cod_emp)values("
    strSQL = strSQL & "'" & sNum & "',"
    strSQL = strSQL & "'" & strCod_Cai & "',"
    strSQL = strSQL & "'" & strCod_Tur & "',"
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & "'" & strIP_Usu & "',"
    strSQL = strSQL & F_GravaDataTempo(HoraServidor) & ","
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'" & cod_For & "', "
    strSQL = strSQL & "'" & cpf & "', "
    strSQL = strSQL & "'" & tipo_ven & "', "
    strSQL = strSQL & "'" & sCodEmp & "') "
    conn.Execute strSQL
    
    F_GeraNovaVenda = sNum
    sIp = F_BuscaUmCampo("ip_ven", TAB_VENDAS, "num_ven='" & sNum & "'")
    If sIp <> strIP_Usu Then
        MsgBox " A venda " & sNum & " já foi gerada para o IP: " & sIp, vbExclamation
        F_GeraNovaVenda = ""
    End If
    
    
    
    Exit Function
trata_Erro:
    MostraErro "Erro na criação da venda, tente novamente!"
End Function
Function F_GeraNovaVendaFuncionario(Optional cod_cli As String = "000") As String
    Dim sNum As String
    Dim sIp As String
    On Error GoTo trata_Erro
    
   
    
    
    sNum = ChaveSQL(TAB_VENDAS, "num_ven", "000000#")
    
    If F_BuscaUmCampo("num_ven", "tab_Vendas", "num_ven='" & sNum & "'") <> "" Then
        F_GeraNovaVendaFuncionario = ""
        Exit Function
        
    End If
    
    strSQL = "insert into " & TAB_VENDAS & "(num_ven,cod_cli,data_ven,abrevenda_usu,cod_for,ip_ven,tipocli_ven)values("
    strSQL = strSQL & "'" & sNum & "',"
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & F_GravaDataTempo(HoraServidor) & ","
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'06'," ' a prazo 1 x
    strSQL = strSQL & "'" & strIP_Usu & "',"
    strSQL = strSQL & "'U')"
    conn.Execute strSQL
    
    F_GeraNovaVendaFuncionario = sNum
    
    F_GeraNovaVendaFuncionario = sNum
    sIp = F_BuscaUmCampo("ip_ven", TAB_VENDAS, "num_ven='" & sNum & "'")
    If sIp <> strIP_Usu Then
        MsgBox " A venda " & sNum & " já foi gerada para o IP: " & sIp, vbExclamation
        F_GeraNovaVendaFuncionario = ""
    End If
    
    
    Exit Function
trata_Erro:
    MostraErro "Erro na criação da venda, tente novamente!"
End Function

Function F_MudaQtdEstoque(cod_pro As String, valor, Optional real As Boolean = True) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    F_MudaQtdEstoque = True
    
    strSQL = "update " & TAB_PRODUTOS & " set "
    If real Then
        strSQL = strSQL & " er_pro=er_pro " & F_GravaDec(valor)
    Else
        strSQL = strSQL & " ef_pro=ef_pro " & F_GravaDec(valor)
    End If
    strSQL = strSQL & " where cod_pro='" & cod_pro & "'"
    
    conn.Execute strSQL
    
    
    Exit Function
trata_Erro:
    F_MudaQtdEstoque = False
End Function
Function F_BuscaItens(num_ven As String, Optional sOrd As String = "seq_ite", Optional sWhere As String = "") As String
    Dim strSQL As String
    strSQL = "select ROW_NUMBER() OVER () row_num,  if(pq_itd>0,'S','N') as pq,nome_mar,pv_pro,"
        
    strSQL = strSQL & "(ti.qtd_ite*ti.unit_ite) as total,tp.er_pro,desconto_ite as desconto ,final_ite as final, "
    
    strSQL = strSQL & " tp.icms_pro,tp.reservado_pro,tp.codfab_pro,nome_cat,ti.*  "
    strSQL = strSQL & " ,(desconto_ite/ (ti.qtd_ite*ti.unit_ite)*100 ) AS desconto_percent "
    
    If bUsaConfeccao Then
        strSQL = strSQL & ",id.* "
    End If
    
    
    
    strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp inner join "
    
    
    strSQL = strSQL & TAB_ITVENDAS & "   as ti on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " left join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
    If bUsaConfeccao Then
        strSQL = strSQL & " left join itvendas_detalhe as id on ti.id=id.id    "
    End If
    
    
    strSQL = strSQL & " inner join " & TAB_CATEGORIAS & " as cat on cat.cod_cat=tp.cod_cat"
    
    strSQL = strSQL & " where  "
    
    strSQL = strSQL & "  ti.num_ven='" & num_ven & "' "
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
    strSQL = strSQL & " from " & TAB_CLIENTES & " as tc, "
    strSQL = strSQL & TAB_VENDAS & " as tv, tab_funcionarios as tf,tab_formaspag as tfo  "
    
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
    strSQL = strSQL & " from " & TAB_CLIENTES & " as tc, "
    strSQL = strSQL & TAB_VENDAS & " as tv, tab_formaspag as tfo  "
    
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
    strSQL = "Update " & TAB_ITVENDAS & " set desconto_ite=desconto_ite+(final_ite/100)*" & F_GravaDec(desconto)
    
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    If queima = False Then
        strSQL = strSQL & " and pq_itd=0 "
    End If
    conn.Execute strSQL
    
    strSQL = "update " & TAB_ITVENDAS & " set final_ite=unit_ite*qtd_ite-desconto_ite "
    strSQL = strSQL & " where num_ven='" & num_ven & "'"
    If queima = False Then
        strSQL = strSQL & " and pq_itd=0 "
    End If
    conn.Execute strSQL
    
    strSQL = "select sum(final_ite) as final, "
    strSQL = strSQL & "sum(desconto_ite) as desconto "
    strSQL = strSQL & "from " & TAB_ITVENDAS & " where num_ven='" & num_ven & "' "
    rPesq.Open strSQL, conn
    
    With rPesq
        strSQL = "update " & TAB_VENDAS & " set desconto_ven=" & F_GravaDec(!desconto) & ","
        strSQL = strSQL & "valor_ven=" & F_GravaDec(novoValor_ven) & ", "
        strSQL = strSQL & "autorizadesconto_ven='" & autoriza_desconto & "' "
        strSQL = strSQL & " where num_ven='" & num_ven & "' "
        conn.Execute strSQL
        .Close
    End With
    
End Sub
Sub P_AtualizaVendaEndereco(num_ven As String, Entrega As String, Endereco As String, bairro As String, _
    Optional telefone As String = "", Optional obs As String = "", Optional comprador As String = "")
    Dim strSQL As String
    
    
    strSQL = "update " & TAB_VENDAS & " set endereco_Ven='" & Endereco & "', "
    strSQL = strSQL & " entrega_ven='" & Entrega & "' "
    'If bairro <> "" Then
        strSQL = strSQL & ",bairro_ven='" & bairro & "'  "
    'End If
    'If telefone <> "" Then
        strSQL = strSQL & ",telefonecli_ven='" & telefone & "'  "
    'End If
    
    'If obs <> "" Then
        strSQL = strSQL & ",obs_ven='" & obs & "' "
    'End If
    strSQL = strSQL & ",comprador_ven='" & comprador & "' "
    
    
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    conn.Execute strSQL

End Sub

Sub P_AtualizaVendaCliente(num_ven As String, cod_cli As String, nomecli_ven As String, _
    Optional telefone As String = "", Optional cpfcli_ven As String = "")
    Dim strSQL As String
    Dim i As Integer
    If cod_cli = "" Or cod_cli = "0" Or cod_cli = "00000" Or cod_cli = "000000" Or cod_cli = "0000000" Then
        If nomecli_ven = "" Then nomecli_ven = "Venda a Consumidor"
    End If
    If cod_cli = "" Then
        cod_cli = "0"
    End If
    strSQL = "update " & TAB_VENDAS & " set cod_cli='" & cod_cli & "' "
    If nomecli_ven <> "" Then
        strSQL = strSQL & ",nomecli_ven='" & nomecli_ven & "'  "
    End If
    If cpfcli_ven <> "" Then
        strSQL = strSQL & ",cpfcli_ven='" & cpfcli_ven & "'  "
    End If
    
    If telefone <> "" Then
        strSQL = strSQL & ",telefonecli_ven='" & telefone & "' "
    End If
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    conn.Execute strSQL, i

End Sub
Sub P_AtualizaValorTotalVenda(num_ven As String, valor As String, desconto As String)
    Dim strSQL As String
    strSQL = "update " & TAB_VENDAS & " set valor_ven=" & F_GravaDec(valor) & ","
    strSQL = strSQL & "desconto_ven=" & F_GravaDec(desconto) & " "
    strSQL = strSQL & " where num_ven='" & num_ven & "'"
    conn.Execute strSQL
    
    
End Sub
Sub P_AtualizaVendaValores(num_ven As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    Dim cDesc As Currency
    strSQL = "select final_ite,desconto_ite from " & TAB_ITVENDAS & " where num_ven='" & num_ven & "' "
    rPesq.Open strSQL, conn
    With rPesq
        Do While Not .EOF
            cTotal = cTotal + !final_ite
            cDesc = cDesc + !Desconto_ite
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "update " & TAB_VENDAS & " set valor_ven=" & F_GravaDec(cTotal) & ","
    strSQL = strSQL & "desconto_ven=" & F_GravaDec(cDesc)
    strSQL = strSQL & " where num_ven='" & num_ven & "'"
    conn.Execute strSQL

End Sub



Sub P_AtualizaFormaPag(num_ven, formaPag As String)
    Dim strSQL As String
    strSQL = "update " & TAB_VENDAS & " set cod_For='" & formaPag & "' "
    
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    conn.Execute strSQL

End Sub

Function F_AlteraStatusVenda(num_ven As String, status As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update " & TAB_VENDAS & " set situacao_ven='" & status & "' "
    
    If status = "2" Then
        strSQL = strSQL & ",fechavenda_usu='" & strCod_Usu & "' "
        strSQL = strSQL & ",data_ven=" & F_GravaDataTempo(Now)
        
        
    End If
    
    
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    conn.Execute strSQL
    
    F_AlteraStatusVenda = True
    
    Exit Function
trata_Erro:
    
End Function

Function F_BuscaItVendas(num_ven As String) As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select ti.* from " & TAB_ITVENDAS & " as ti where num_ven='" & num_ven & "' "
    rPesq.Open strSQL, conn
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
Function CancelaItem(num_ven As String, seq_ite As String) As Boolean
    Dim strSQL As String
    Dim i As Integer
    On Error GoTo trata_Erro
    Dim qtd As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select seq_ite,cod_pro,qtd_ite, cod_emp from " & TAB_ITVENDAS & " where num_ven='" & num_ven & "' "
    If LCase(seq_ite) <> "todos" Then
        strSQL = strSQL & " and seq_ite=" & seq_ite
    End If
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            strSQL = "delete from " & TAB_ITVENDAS & " where num_ven='" & num_ven & "'"
            strSQL = strSQL & " and seq_ite='" & !seq_ite & "'"
            conn.Execute strSQL, i
            
            strSQL = "update " & TAB_PRODUTOS & " set reservado_pro=reservado_pro-" & str(!qtd_ite)
            strSQL = strSQL & " where cod_pro='" & rPesq!cod_pro & "'"
            conn.Execute strSQL
            strSQL = "update estoque set reservado=reservado-" & str(!qtd_ite)
            strSQL = strSQL & " where cod_pro='" & rPesq!cod_pro & "'"
            strSQL = strSQL & " and cod_emp='" & rPesq!cod_emp & "'"
            conn.Execute strSQL
            
            'P_GravaItemCancelado
            
            .MoveNext
        Loop
        End With
    
    CancelaItem = True
    Exit Function
trata_Erro:
    CancelaItem = False
End Function


Function F_AdicionaItem(num_ven As String, prod As Produto, qtd As String, desconto As String, _
    Optional final As String = "0", Optional pq_itd As String = "0", Optional cod_emp As String = "01", _
    Optional atacado_ite As String = "N", Optional detalhamento As String = "") As String
    Dim sSeq As String
    Dim id As Long
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    Dim desc_apl As Currency
    Dim cTotal As Currency
    On Error GoTo trata_Erro
    
    If final = "" Or final = "0" Then
        final = (CCur(qtd) * CCur(prod.pv_pro)) - CCur(desconto)
    End If
    
    cTotal = CCur(qtd) * CCur(prod.pv_pro)
    
   ' conn.BeginTrans
   
    If cTotal > 0 Then
        desc_apl = CCur(desconto) / (CCur(qtd) * CCur(prod.pv_pro)) * 100
    End If
   
   
    
    sSeq = MaxKeySQL(TAB_ITVENDAS, "seq_ite", "num_ven='" & num_ven & "'", "")
    strSQL = "insert into " & TAB_ITVENDAS & "(num_ven,seq_ite,cod_pro,qtd_ite,unit_ite,desc_ite,"
    strSQL = strSQL & " desconto_ite,final_ite,pq_itd,cod_emp,ean_ite,atacado_ite,desconto_aplicado"
    If detalhamento <> "" Then
        strSQL = strSQL & ",detalhamento"
    End If
    
    strSQL = strSQL & " )values("
    
    
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & "'" & prod.cod_pro & "',"
    
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & F_GravaDec(prod.pv_pro) & ","
    If prod.codfab_pro <> "" Then
        strSQL = strSQL & "'" & prod.codfab_pro & ": " & prod.nome_pro & "',"
    Else
        strSQL = strSQL & "'" & prod.nome_pro & "',"
    End If
    strSQL = strSQL & F_GravaDec(desconto) & ","
    strSQL = strSQL & F_GravaDec(final) & ","
    strSQL = strSQL & F_GravaDec(pq_itd) & ","
    strSQL = strSQL & "'" & cod_emp & "', "
    strSQL = strSQL & "'" & prod.ean_pro & "', "
    strSQL = strSQL & "'" & prod.Atacado & "',"
    strSQL = strSQL & F_GravaDec(desc_apl)
    If detalhamento <> "" Then
        strSQL = strSQL & ",'" & detalhamento & "'"
    End If
    
    strSQL = strSQL & ") "
    conn.Execute strSQL
    
    strSQL = "select max(id) as id  from itvendas"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        id = !id
        .Close
    End With
    
    
        
    strSQL = "update " & TAB_PRODUTOS & " set reservado_pro=reservado_pro+" & str(qtd)
    strSQL = strSQL & " where cod_pro='" & prod.cod_pro & "';"
    conn.Execute strSQL
    strSQL = "update estoque set reservado=reservado+" & str(qtd)
    strSQL = strSQL & " where cod_pro='" & prod.cod_pro & "'"
    strSQL = strSQL & " and cod_emp='" & cod_emp & "';"
    conn.Execute strSQL
    strSQL = "update " & TAB_VENDAS & " set cod_emp='" & cod_emp & "'"
    strSQL = strSQL & " where num_ven='" & num_ven & "';"
    conn.Execute strSQL
    
    If bUsaConfeccao Then
        If prod.detalhe.bordado <> "" Or prod.detalhe.cor <> "" Or prod.detalhe.silk <> "" Or prod.detalhe.detalhamento <> "" Then
            strSQL = "insert into itvendas_detalhe(id,cor,silk,bordado,detalhamento)values("
            strSQL = strSQL & id & ", "
            strSQL = strSQL & "'" & prod.detalhe.cor & "',"
            strSQL = strSQL & "'" & prod.detalhe.silk & "',"
            strSQL = strSQL & "'" & prod.detalhe.bordado & "',"
            strSQL = strSQL & "'" & prod.detalhe.detalhamento & "')"
            conn.Execute strSQL
        End If
    End If
    'conn.CommitTrans
    
    
    
    
    F_AdicionaItem = sSeq
    Exit Function
trata_Erro:
MsgBox Err.Description
End Function

Function F_GravaCrediarioVen(num_ven As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tab_crediarioven(num_ven,status_cre,data_cre)values('" & num_ven & "',"
    strSQL = strSQL & "'1',"
    strSQL = strSQL & F_GravaDataTempo(HoraServidor) & ")"
    conn.Execute strSQL
    
    
    strSQL = "update " & TAB_VENDAS & " set crediario_ven=1"
            
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    conn.Execute strSQL
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
    strSQL = strSQL & "fechamento_cre=" & F_GravaDataTempo(HoraServidor)
    strSQL = strSQL & " where num_ven='" & num_ven & "'"
    conn.Execute strSQL
    
    conn.Execute strSQL
    F_AlteraCrediarioVen = True
    Exit Function
trata_Erro:
    
End Function


Sub AtualizaDesconto(num_ven As String, desconto As String)
    
    Dim strSQL As String
    desconto = CCur(desconto / 100)
    
    
    
    strSQL = "update " & TAB_ITVENDAS & " set desconto_ite=(qtd_ite*unit_ite)*" & F_GravaDec(desconto) & ","
    strSQL = strSQL & "final_ite=(qtd_ite*unit_ite)-((qtd_ite*unit_ite)*" & F_GravaDec(desconto) & "),"
    strSQL = strSQL & "desconto_aplicado=((qtd_ite*unit_ite-final_ite)/(qtd_ite*unit_ite)*100),"
    strSQL = strSQL & "pq_itd=0 "
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    conn.Execute strSQL
 
    
    P_AtualizaVendaValores num_ven
    verifica_atacado num_ven
End Sub



Sub P_CancelaDescontoVenda(num_ven As String, Optional desconto As String = "0")
    Dim strSQL As String
    desconto = CCur(desconto / 100)
    strSQL = "update " & TAB_ITVENDAS & " set desconto_ite=(qtd_ite*unit_ite)*" & F_GravaDec(desconto) & ","
    strSQL = strSQL & "final_ite=(qtd_ite*unit_ite)-((qtd_ite*unit_ite)*" & F_GravaDec(desconto) & "),"
    strSQL = strSQL & "pq_itd=0 "
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    conn.Execute strSQL
    strSQL = "update " & TAB_VENDAS & " set autorizadesconto_ven='' where num_ven='" & num_ven & "'"
    conn.Execute strSQL
    
    P_AtualizaVendaValores num_ven
End Sub
Function F_AdicionaVenFec(num_ven As String, valor_ven As String, _
    cod_For As String, cod_cai As String, cod_tur As String, esp_cai As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    If cod_For = "" Then
        cod_For = "01"
    End If
    strSQL = "insert into " & TAB_VENFEC & "(num_ven,valor_ven,cod_for,cod_cai,cod_tur,esp_cai)values("
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & F_GravaDec(valor_ven) & ","
    strSQL = strSQL & "'" & cod_For & "',"
    strSQL = strSQL & "'" & cod_cai & "',"
    strSQL = strSQL & "'" & cod_tur & "',"
    strSQL = strSQL & "'" & esp_cai & "')"
    conn.Execute strSQL
    F_AdicionaVenFec = True
    strSQL = "update " & TAB_VENFEC & " set "
    If cod_For = "01" Or cod_For = "" Then
        strSQL = strSQL & "avista_ven=" & F_GravaDec(valor_ven)
    ElseIf cod_For = "09" Or cod_For = "10" Then
        strSQL = strSQL & "finan_ven=" & F_GravaDec(valor_ven)
    Else
        strSQL = strSQL & "aprazo_ven=" & F_GravaDec(valor_ven)
    End If
    strSQL = strSQL & " where num_ven='" & num_ven & "'"
    strSQL = strSQL & " and cod_for='" & cod_For & "' "
    conn.Execute strSQL
Exit Function
trata_Erro:
    
End Function
    


Sub P_AutorizaDescontoGerencial(num_ven As String, desconto As String)
    Dim strSQL As String
    strSQL = "update " & TAB_ITVENDAS & " set desconto_ite=desconto_ite+(" & F_GravaDec(desconto)
    strSQL = strSQL & ""
End Sub
Sub P_GravaVendaAberta(num_ven As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    'If bUsaCaixaSeparado = True Then
        P_DeletaVendaAberta num_ven
        strSQL = "insert into vendaaberta(num_ven,ip_ven)values("
        strSQL = strSQL & "'" & num_ven & "',"
        strSQL = strSQL & "'" & strIP_Usu & "') "
        conn.Execute strSQL
    'End If
    Exit Sub
trata_Erro:
    
End Sub
Sub P_DeletaVendaAberta(num_ven As String)
    On Error Resume Next
    
    strSQL = "delete from vendaaberta where num_ven='" & num_ven & "'"
    
    conn.Execute strSQL
    
End Sub
Function F_VerificaVendaAberta(num_ven As String) As String
    Dim strSQL As String
    strSQL = F_BuscaUmCampo("ip_ven", "vendaaberta", "num_ven='" & num_ven & "'")
    If strSQL <> strIP_Usu Then
    
        F_VerificaVendaAberta = strSQL
    End If
End Function
Function F_CancelaVendaAposCupom(sNum_ven As String) As Boolean
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    Dim sCodCai As String
    'varre toda a venda voltando o estoque real dos itens depois muda para situacao 5
    F_CancelaVendaAposCupom = False
    
    
    If F_BuscaUmCampo("status_ven", TAB_VENDAS, "num_ven='" & sNum_ven & "'") = "3" Then
        MsgBox "Não foi possível cancelar os lançamentos da venda, pois ela já se encontra cancelada!", vbExclamation
        F_CancelaVendaAposCupom = False
        Exit Function
    End If
    sCodCai = F_BuscaUmCampo("Cod_Cai", TAB_VENFEC, "num_ven='" & sNum_ven & "'")
    
    'P_cancelaCupon F_BuscaUmCampo("seq_cup", "cupons", "num_ven='" & sNum_ven & "' and cancelado_cup='N'")
    strSQL = "update tab_saidas set serie_sai='CANC' where num_ven='" & sNum_ven & "'"
    conn.Execute strSQL
    
    
    strSQL = "select * from " & TAB_ITVENDAS & " where num_ven='" & sNum_ven & "'  "
    rPesq.Open strSQL, conn
    With rPesq
        Do While Not .EOF
        
        
            
        
        
        
        
            strSQL = "update " & TAB_PRODUTOS & " set er_pro=er_pro + " & F_GravaDec(CStr(!qtd_ite))
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            
            conn.Execute strSQL
        
        
            strSQL = "update estoque set er=er + " & F_GravaDec(CStr(!qtd_ite))
            strSQL = strSQL & " where cod_emp='" & !cod_emp & "' "
            strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
            conn.Execute strSQL
        
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "select * from " & TAB_VENFEC & " where num_ven='" & sNum_ven & "' "
    rPesq.Open strSQL, conn
    With rPesq
        Do While Not .EOF
            
            strSQL = F_GravaCai("D", Date, CStr(!valor_ven), "Venda " & !num_ven & " Cancelada", !num_ven, _
                F_BuscaUmCampo("tipo_cai", "tab_formaspag", "cod_for='" & !cod_For & "'"), "01", _
                0, 0, sCodCai, strCod_Tur, "000", "S")
            
            
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "select * from  " & TAB_DOC & "  where num_ven='" & sNum_ven & "' and esp_doc='V' "
                rPesq.Open strSQL, conn
                With rPesq
                    Do While Not .EOF
                        strSQL = "delete from  " & TAB_ITDOC & "  where seq_doc='" & !seq_doc & "'"
                        conn.Execute strSQL
                        strSQL = "delete from tab_cheques where seq_doc='" & !seq_doc & "'"
                        conn.Execute strSQL
                        strSQL = "delete from  " & TAB_DOC & "  where seq_doc='" & !seq_doc & "'"
                        conn.Execute strSQL
                        .MoveNext
                    Loop
                End With
    
    strSQL = "update " & TAB_VENDAS & " set situacao_ven=5, "
    strSQL = strSQL & " cancelavenda_usu='" & strCod_Usu & "' "
    strSQL = strSQL & " where num_ven='" & sNum_ven & "' "
    conn.Execute strSQL '5 = venda cancelada
    F_CancelaVendaAposCupom = True
    Exit Function
trata_Erro:
    MostraErro
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
    strSQL = strSQL & F_GravaDataTempo(HoraServidor) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & cliente & "',"
    strSQL = strSQL & "'" & vendedor & "',"
    strSQL = strSQL & "'" & obs & "',"
    
    strSQL = strSQL & F_GravaDec(valor_in) & ","
    strSQL = strSQL & F_GravaDec(valor_out) & ","
    strSQL = strSQL & "'" & telefone & "',"
    strSQL = strSQL & F_GravaDec(desconto) & ")"
    conn.Execute strSQL



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
    conn.Execute strSQL
    Exit Sub
trata_Erro:
    MostraErro
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
    conn.Execute strSQL
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
    conn.Execute strSQL
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
'        P_ImprX obj, "" & F_BuscaUmCampo("guerra_fun", tab_funcionarios, "cod_fun='" & !fechavenda_usu & "'"), 45
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
Sub AtualizaVenda(num_ven As String, valor_ven As String, desconto_ven As String, nomecli_ven As String, telefonecli_ven, _
    comprador_ven As String, _
    cod_cli As String, cod_cai As String, abrevenda_usu As String, obs_ven As String, email_ven As String)
    Dim strSQL As String
    
    strSQL = "update " & TAB_VENDAS & " set valor_ven =" & F_GravaDec(valor_ven) & ","
    strSQL = strSQL & "final_ven=" & F_GravaDec(valor_ven - desconto_ven) & ","
    strSQL = strSQL & "desconto_ven=" & F_GravaDec(desconto_ven) & ","
    strSQL = strSQL & "nomecli_ven= '" & nomecli_ven & "', "
    strSQL = strSQL & "telefonecli_ven= '" & telefonecli_ven & "', "
    strSQL = strSQL & " comprador_ven= '" & comprador_ven & "', "
    strSQL = strSQL & " cod_cli= '" & cod_cli & "', "
    strSQL = strSQL & " cod_cai='" & cod_cai & "',  "
    strSQL = strSQL & " abrevenda_usu='" & abrevenda_usu & "', "
    strSQL = strSQL & " email_ven='" & email_ven & "', "
    strSQL = strSQL & " obs_VEN='" & obs_ven & "' "
    strSQL = strSQL & " where num_Ven='" & num_ven & "' "
    conn.Execute strSQL
    
End Sub
Function GravaVendaCompleta(num_ven As String, cod_cli As String, data_ven As String, dataefet_ven As String, _
ip_ven As String, valor_ven As String, abrevenda_usu As String, fechavenda_usu As String, cod_cai As String, _
situacao_ven As String, desconto_ven As String, cod_For As String, nomecli_ven As String, cancelavenda_usu As String, _
entrega_ven As String, endereco_ven As String, bairro_ven As String, fone_ven As String, crediario_ven As String, _
status_ven As String, autorizadesconto_ven As String, tipocli_ven As String, telefonecli_ven As String, cpfcli_ven As String, _
obs_ven As String, final_ven As String, seqefet_ven As String) As Boolean

    Dim strSQL As String
    On Error GoTo trata_Erro

    If seqefet_ven = "" Then seqefet_ven = "0"

    strSQL = "replace into " & TAB_VENDAS & "(num_ven,cod_cli,data_ven,dataefet_ven,ip_ven,valor_ven,"
    strSQL = strSQL & "abrevenda_usu,fechavenda_usu,cod_cai,situacao_ven,desconto_ven,cod_for,nomecli_ven,"
    strSQL = strSQL & "cancelavenda_usu,entrega_ven,endereco_ven,bairro_ven,fone_ven,crediario_ven,"
    strSQL = strSQL & "status_ven,autorizadesconto_ven,tipocli_ven,telefonecli_ven,cpfcli_ven,"
    strSQL = strSQL & "obs_ven,final_ven,seqefet_ven)values("
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & F_GravaDataTempo(data_ven) & ","
    strSQL = strSQL & F_GravaData(dataefet_ven) & ","
    strSQL = strSQL & "'" & ip_ven & "',"
    strSQL = strSQL & F_GravaDec(valor_ven) & ","
    strSQL = strSQL & "'" & abrevenda_usu & "',"
    strSQL = strSQL & "'" & fechavenda_usu & "',"
    strSQL = strSQL & "'" & cod_cai & "',"
    strSQL = strSQL & "'" & situacao_ven & "',"
    strSQL = strSQL & F_GravaDec(desconto_ven) & ","
    strSQL = strSQL & "'" & cod_For & "',"
    strSQL = strSQL & "'" & nomecli_ven & "',"
    strSQL = strSQL & "'" & cancelavenda_usu & "',"
    strSQL = strSQL & "'" & entrega_ven & "',"
    strSQL = strSQL & "'" & endereco_ven & "',"
    strSQL = strSQL & "'" & bairro_ven & "',"
    strSQL = strSQL & "'" & fone_ven & "',"
    strSQL = strSQL & "'" & crediario_ven & "',"
    strSQL = strSQL & "'" & status_ven & "',"
    strSQL = strSQL & "'" & autorizadesconto_ven & "',"
    strSQL = strSQL & "'" & tipocli_ven & "',"
    strSQL = strSQL & "'" & telefonecli_ven & "',"
    strSQL = strSQL & "'" & cpfcli_ven & "',"
    strSQL = strSQL & "'" & obs_ven & "',"
    strSQL = strSQL & F_GravaDec(final_ven) & ","
    strSQL = strSQL & "'" & seqefet_ven & "')"
    conn.Execute strSQL
    GravaVendaCompleta = True

    Exit Function
trata_Erro:
'P_MostraErro
End Function


Function NovaVenda() As String
    Dim cpf As String
    Dim strSQL As String
    Dim c As caixa_logado
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Function
    End If
    If bUsaVendaSupermercado Then
        If Not F_IsCaixa Then
            MsgBox "Este computador não é um caixa, portanto não pode efetuar vendas!", vbExclamation
            Exit Function
        End If
        
        If bUsaVendaSupermercado = True Then
            cpf = InformaCPFClienteVenda
        End If
        
        c = Caixa.BuscaCaixaLogado(strCod_Cai)
        If c.status = "F" Then
            MsgBox "O Caixa encontra-se fechado!", vbExclamation
            Exit Function
            
        End If
        
    End If

    If strFormaPagamentoPadrao = "" Then
        strFormaPagamentoPadrao = "01"
    End If
    
    strSQL = F_GeraNovaVenda("000000", strFormaPagamentoPadrao, cpf)
    
    If strSQL = "" Then
        MsgBox "erro! A venda não foi gerada!", vbExclamation
        Exit Function
    End If
    
    
    
    
    If bUsaVendaSupermercado Then
        frmVendaSup.tag = strSQL
        frmVendaSup.Show 1
    Else
        frmVenda.tag = strSQL
        
        frmVenda.Atacado = True
        frmVenda.Show 0, mdiFrm
    End If


    NovaVenda = strSQL

End Function
Sub editaVenda(num_ven As String)
    Dim strSQL As String
    
    
    
    strSQL = "select * from " & TAB_VENDAS & " where num_ven='" & num_ven & "' "
    
    
    
    If bUsaVendaSupermercado = True Then
    
    
        frmVendaSup.tag = num_ven
        
        
        frmVendaSup.Show 1 ', mdiFrm
    Else
        frmVenda.tag = num_ven
        
        
        frmVenda.Show 1
    End If
    
    P_Status ""
    
End Sub

Function BuscaClienteRapidoVenda(codigo As String, cpf As String, nome As String) As ClienteVenda
    frmClienteVenda.txtCPF.Text = cpf
    frmClienteVenda.txtNome.Text = nome
    frmClienteVenda.txtNome.tag = codigo
    frmClienteVenda.Show 1
    If frmClienteVenda.tag <> "" Then
        BuscaClienteRapidoVenda.cpfcnpj = frmClienteVenda.txtCPF.Text
        BuscaClienteRapidoVenda.cod_cli = frmClienteVenda.txtCPF.tag
        BuscaClienteRapidoVenda.nomecli_ven = frmClienteVenda.txtNome.Text
        
        Unload frmClienteVenda
    End If
    
    
End Function
Function InformaCPFClienteVenda() As String
   ' frmCPFVenda.Show 1
   ' InformaCPFClienteVenda = frmCPFVenda.txtCPF.Text
   ' Unload frmCPFVenda
    
End Function
Function CopiaTabelaVendasPorAno(ano As String) As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim ini As String
    Dim final As String
    If ano = Year(Date) Then
        Exit Function
    End If
    
    TAB_VENDAS = "tab_vendas"
    
    strSQL = "select min(num_ven) as ini, max(num_ven) as final from " & TAB_VENDAS
    strSQL = strSQL & " where data_ven>='" & ano & "-01-01 00:00'  "
    strSQL = strSQL & " and data_ven<='" & ano & "-12-31 23:59'"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            ini = rPesq!ini
            final = rPesq!final
        Else
            Exit Function
            .Close
        End If
    End With
    
    
    strSQL = "create table   " & TAB_VENDAS & "_" & ano
    strSQL = strSQL & " select * from  " & TAB_VENDAS & " where num_ven>='" & ini & "' and num_ven<='" & final & "' "
    conn.Execute strSQL
    
    strSQL = "create table   " & TAB_ITVENDAS & "_" & ano
    strSQL = strSQL & " select * from  " & TAB_ITVENDAS & " where num_ven>='" & ini & "' and num_ven<='" & final & "' "
    conn.Execute strSQL
    
    strSQL = "create table   " & TAB_VENFEC & "_" & ano
    strSQL = strSQL & " select * from  " & TAB_VENFEC & " where num_ven>='" & ini & "' and num_ven<='" & final & "' "
    conn.Execute strSQL
    
    strSQL = "create table   " & TAB_VENFIN & "_" & ano
    strSQL = strSQL & " select * from  " & TAB_VENFIN & " where num_ven>='" & ini & "' and num_ven<='" & final & "' "
    conn.Execute strSQL
    
    
    
    
    
    
    strSQL = "delete from " & TAB_VENDAS & "  where num_ven>='" & ini & "' and num_ven<='" & final & "' "
    strSQL = "delete from " & TAB_ITVENDAS & "  where num_ven>='" & ini & "' and num_ven<='" & final & "' "
    strSQL = "delete from " & TAB_VENFIN & "  where num_ven>='" & ini & "' and num_ven<='" & final & "' "
    strSQL = "delete from " & TAB_VENFEC & "  where num_ven>='" & ini & "' and num_ven<='" & final & "' "
    
    
    


    

End Function
Sub AlteraStatusVen(num_ven As String, status_ven As String)
    Dim strSQL As String
    strSQL = "update " & TAB_VENDAS & " set status_ven='" & status_ven & "' "
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    conn.Execute strSQL
    
    
End Sub
Sub ImprimeComprovanteNaoFiscal(impressora As Object, num_ven As String, Optional bViz As Boolean = False)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim crlf_ As String
    crlf_ = CrLf
    On Error GoTo trata_Erro
    'Dim impressora As Object
    Dim printpadrao As Object
    Dim iPrint As Integer
    iPrint = BuscaImpressoraPadrao
    Set Printer = Printers(ImpressoraPadraoCupomNFCe.Indice)
    
    
    
    If bUsaVendaSupermercado = True Then
        If isDesignTime = True Then
            bViz = True
        Else
            bViz = False
        End If
        bViz = False
    End If
    
    If bViz = False Then
        Set impressora = Printer
    Else
    
        Set impressora = frmVisualiza.Pic
    End If
    
    
    
    
    Dim Centra As Integer
    Dim cTotal As Currency
    Dim valor_ven As Currency
    Dim nome_Cli As String
    Centra = 60
    
    strSQL = "select tv.num_ven,tv.valor_ven,tv.data_ven,tv.abrevenda_usu,tv.fechavenda_usu,tv.valor_ven,tv.desconto_ven,"
    strSQL = strSQL & "tv.nomecli_ven,tv.entrega_ven,tv.endereco_ven,ti.* from " & TAB_VENDAS & " as tv inner join " & TAB_ITVENDAS & " as ti "
    strSQL = strSQL & " on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " where tv.num_ven='" & num_ven & "' "
    strSQL = strSQL & " and tv.situacao_ven='4' "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Venda não encontrada em situação efetivada.", vbExclamation
            
            .Close
            Exit Sub
        End If
        
        
        imprimeLogo impressora
       ' impressora.FontName = "arial"
        If TypeName(impressora) = "Printer" Then
            impressora.Print
        Else
            impressora.Print
        End If
        P_CentraImpressao impressora, String(40, "*")
        P_CentraImpressao impressora, Trim(strNome_Emp)
        P_CentraImpressao impressora, Trim(strEndereco_Emp)
        P_CentraImpressao impressora, "FONE:" & strFone_Emp
        P_CentraImpressao impressora, String(40, "*")
        'P_ImprX impressora, F_Centra(strNome_Emp, 90), 0, False
        'P_ImprX impressora, F_Centra(strEndereco_Emp, 90), 0, False
        'P_ImprX impressora, F_Centra("FONE:" & strFone_Emp, 90), 0, False
        'P_ImprX impressora, F_Centra("*", Centra, "*"), 0, False
        
        P_line1 impressora
        impressora.FontBold = True
        impressora.Print
        impressora.Print
        'P_PrintXY impressora, F_Centra("ROMANEIO DE ENTREGA", 50, "*"), 0, impressora.CurrentY, 2
        P_CentraImpressao impressora, "ROMANEIO DE ENTREGA"
        'impressora.Print

        impressora.FontBold = False
        impressora.Print
        impressora.Print
        P_ImprX impressora, "VENDA:" & !num_ven & " DATA: " & !data_ven & " Op:" & !abrevenda_usu, 0, False
        impressora.FontBold = True
                P_ImprQuebra "CLIENTE:" & !nomecli_ven, impressora, 0, impressora.ScaleWidth
                nome_Cli = !nomecli_ven
        impressora.Print
        P_ImprX impressora, "ENDEREÇO:" & !endereco_ven, 0, False
        impressora.FontBold = False
        impressora.Print
        P_line1 impressora
        impressora.Print
        P_line1 impressora
        valor_ven = !valor_ven
        P_line1 impressora, vbDot
        P_ImprX impressora, "Item Descrição/Qtd x Unitário = Total ", 0, False
        P_line1 impressora, vbDot
        Do While Not .EOF
             
            P_ImprX impressora, Format(.AbsolutePosition, "##0") & ". " & !desc_ite, 1, False
            P_ImprX impressora, FormatVal(!qtd_ite, "###,##0.000") & " X " & Format(!unit_ite) & " = " & Format(!qtd_ite * !unit_ite), 22, False
           
            cTotal = cTotal + (!unit_ite * !qtd_ite)
            
            .MoveNext
        Loop
        impressora.Print
        P_line1 impressora
        impressora.Print
        impressora.FontBold = True
        P_ImprX impressora, "Itens: " & .RecordCount, 0, False
        P_ImprX impressora, "Soma: " & FormatVal(cTotal), 0, False
        P_ImprX impressora, "Valor Final: " & FormatVal(valor_ven), 0, False
        P_line1 impressora, vbInsideSolid
        impressora.FontBold = False
        impressora.Print
        impressora.fontsize = 10
        
        
        'P_ImprX impressora, " MUITO MAIS ATACADO ", 0, False
        
        
        
        
        P_CentraImpressao impressora, "______________________________"
        P_CentraImpressao impressora, nome_Cli
        
        'P_PrintXY impressora, F_Centra("______________________________", 40, "_"), 0, impressora.CurrentY, 2
        'P_PrintXY impressora, F_Centra(nome_Cli, 40, "*"), 0, impressora.CurrentY, 2
        impressora.Print
        
        
        P_CentraImpressao impressora, F_Centra("NÃO É DOCUMENTO FISCAL", 50, "*")
        P_CentraImpressao impressora, F_Centra("NÃO É VÁLIDO COMO GARANTIA DE MERCADORIA", 40, "*")
        P_ImprX impressora, F_Centra("Obrigado! Volte sempre!", 70, "*"), 0, False
        impressora.Print
        impressora.Print Now
        
    End With
    
    
    
    
'    If TypeName(impressora) = "PictureBox" Then
'        frmVisualiza.Show 1
'        Unload frmVisualiza
'    Else
        If bViz Then
            frmVisualiza.Show 1
            Unload frmVisualiza
        Else
            Printer.EndDoc
        End If
        
'    End If
    
        
    Set Printer = Printers(iPrint)
    
Exit Sub
trata_Erro:
MostraErro


End Sub
Sub EnviarEmailVenda(num_ven As String)
    
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
    Arquivo = "Venda" & num_ven & ".pdf"
    
    c = F_BuscaClienteVenda(num_ven)
    
    emailCliente = c.email_cli
    emailCliente = InputBox("Confirme o e-mail a ser enviado", "Email", emailCliente)
    If emailCliente = "" Then
        Exit Sub
    End If
    email = BuscaMailNome("Vendas")
    If email.id = "" Then
        MsgBox "Não foi possível encontrar a configuração de Email.", vbExclamation
        Exit Sub
    End If
    
    ApontarPrinterToPDF Arquivo, , False
    imprimeVenda Printer, num_ven
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
Sub imprimeLogo(obj As Object, Optional X As Single, Optional Y As Single, Optional wid As Single)
    Dim path As String
    Dim w As Single
    Dim d As Single
    Dim p As IPictureDisp
    On Error GoTo trata_Erro

   
    
   
        
    path = strCaminho_Logo
    Set p = LoadPicture(path)
    
    w = p.width
    d = wid / w
    
    obj.PaintPicture LoadPicture(path), X, Y, p.width * d, p.height * d
    obj.Print
    obj.Print
    Set p = Nothing
    Exit Sub
trata_Erro:
    'MsgBox Err.Description
End Sub
Sub EmiteUltRomaneio(strCodCai As String)
    Dim sNumVen As String
    sNumVen = F_BuscaUmCampo("max(num_ven)", "tab_vendas", "cod_cai='" & strCod_Cai & "' and situacao_ven='4'")
    
    If sNumVen <> "" Then
        ImprimeComprovanteNaoFiscal Printer, sNumVen
    End If
End Sub


Sub EditaVenFin(num_ven As String)
    frmLanVenFin.tag = num_ven
    frmLanVenFin.Show 1
    

End Sub
Function AvisaVendasAbertasPorUsuario(dias As Integer)
    Dim strSQL As String
    
    strSQL = F_BuscaUmCampo("count(*)", TAB_VENDAS, "situacao_ven=1 and abrevenda_usu='" & _
        strCod_Usu & "' and data_ven<=" & F_GravaDataTempo(DateAdd("D", -1, Date)))
    If strSQL <> "" And strSQL <> "0" Then
        MsgBox "Você possui " & strSQL & " vendas abertas. Por favor, verifique se elas não precisam ser canceladas.", vbExclamation
    End If
    
    
End Function
Sub dividir_venda_EmpresasMultiplas(num_ven As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rVendas As New ADODB.Recordset
    Dim rItens As New ADODB.Recordset
    Dim sNum_ven As String
    Dim id As String
   
    strSQL = "select distinct num_ven,cod_emp from " & TAB_ITVENDAS
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    'strSQL = strSQL & " and status_ven<3"
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount <= 1 Then
        MsgBox "Não é possível dividir essa venda. Ela não possui múltiplas empresas!", vbExclamation
        Exit Sub
    Else
        If MsgBox("Foram encontrados itens de " & rPesq.RecordCount & " Empresas diferentes. Confirma divisão em " & _
            rPesq.RecordCount & " vendas?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    
    strSQL = "select * from " & TAB_VENDAS & " where num_ven='" & num_ven & "'"
    abreObjetoAdo rVendas, strSQL
    If rVendas!situacao_ven > 2 Then
        MsgBox "A venda não pode ser dividida!", vbExclamation
        Exit Sub
    End If
    With rPesq
        '.MoveNext 'mover para o segundo registro
        Do While Not .EOF
            If .AbsolutePosition = 1 Then
                sNum_ven = !num_ven
                
            Else
                sNum_ven = vendas.F_GeraNovaVenda(rVendas!cod_cli, rVendas!cod_For, rVendas!cpfcli_ven, "V", !cod_emp)
            End If
            
            
            strSQL = "update " & TAB_VENDAS & " set valor_ven=(select sum(final_ite) from " & TAB_ITVENDAS & " as i where i.num_ven='" & !num_ven & "' "
            strSQL = strSQL & " and cod_emp='" & !cod_emp & "'), "
            strSQL = strSQL & " obs_ven='" & rVendas!obs_ven & "',"
            strSQL = strSQL & " nomecli_ven='" & rVendas!nomecli_ven & "',"
            strSQL = strSQL & " comprador_ven='" & rVendas!comprador_ven & "',"
            strSQL = strSQL & " entrega_ven='" & rVendas!entrega_ven & "',"
            strSQL = strSQL & " abrevenda_usu='" & rVendas!abrevenda_usu & "',"
            strSQL = strSQL & " endereco_ven='" & rVendas!endereco_ven & "',"
            strSQL = strSQL & " bairro_ven='" & rVendas!bairro_ven & "',"
            strSQL = strSQL & " entrega_ven='" & rVendas!entrega_ven & "',"
            strSQL = strSQL & " telefonecli_ven='" & rVendas!telefonecli_ven & "',"
            strSQL = strSQL & " venda_principal='" & num_ven & "',"
            strSQL = strSQL & " obs_ven='" & rVendas!obs_ven & "',"
            strSQL = strSQL & " email_ven='" & rVendas!email_ven & "'"
            strSQL = strSQL & " where num_ven='" & sNum_ven & "'"
            conn.Execute strSQL
            
            strSQL = "update " & TAB_ITVENDAS & " set num_ven='" & sNum_ven & "' "
            strSQL = strSQL & " where num_ven='" & !num_ven & "'"
            strSQL = strSQL & " and cod_emp='" & !cod_emp & "'"
            conn.Execute strSQL
            
            
            
              
                
        
            
            
            .MoveNext
        Loop
    
    End With
    
    
    
End Sub

