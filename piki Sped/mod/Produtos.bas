Attribute VB_Name = "Produtos"
Type produtoAtacado
    id As String
    codigo As String
    qtd As String
    preco As String
End Type
Public Enum TTipoMov
    EntradaSped = 1
    saidaSped = 2
End Enum

Type ProdutoConversao
    id As String
    cod_mar As String
    cod_pro As String
    referencia As String
    
End Type


Type ProdutoSaida
    cod_pro As String
    codbar_pro As String
    nome_pro As String
    ref_pro As String
    pv_pro As Currency
    pp_pro As Currency
    
    
    sub As subsec
'    est As Estoque
    cr_pro As Currency
    cf_pro As Currency
    cm_pro As Currency
    
    s_pro As String
    lucro_pro As Currency
    
    icms As Currency
    ecf_icms As String
    fracionado_pro As String
    complemento_pro As String
    pfab_pro As Currency
    nome_fab As String
    cod_fab As String
    bonificado_pro As String
    pa_pro As Currency
    pq_pro As Currency 'pra o caso do produto ser queima
    multiplo As String
End Type

Type Produto
    cod_pro As String
    codbarra_pro As String
    nome_pro As String
    pq_pro As Currency
    r_pro As Currency
    ean_pro As String
    prom_pro As String
    cf_pro As Currency
    cr_pro As Currency
    cod_cat As String
    cod_mar As String
    nome_cat As String
    nome_mar As String
    qtd_pro As Currency
    preco As Currency
    ipi As Currency
    s_pro As String
    lucro_pro As Currency
    codfab_pro As String
    icms As Currency
    ecf_icms As String
    balanca_pro As String
    ncm_pro As String
    totalbalanca_pro As Currency
    adoAtac As ADODB.Recordset
    cod_sec As String
    nome_sec As String
    unidade_pro As String
    cod_sub As String
    nome_sub As String
    cod_gen As String
    nome_gen As String
    qtd As String
    pv_pro As String
    pa_pro As String
    er_pro As String
    ef_pro As String
    cod_emp As String
    fracionado_pro As String
    sub As subsec
    cat As categoria.categoria
    marc As Marca.Marca
    Atacado As String
    controlaestoque_pro As String
    estoque_total As String
    empresa As String
    detalhe As detalhamento_uniforme
    multiplo_pro As String
    
End Type
Type subsecao
    id As String
    nome As String
    id_secao As String
    nome_secao As String
End Type
Type InfoProd
    cod_pro As String
    ult_custo As String
    nota As String
    cod_For As String
    nome_For As String
    data As String
    qtd As Integer
    unit As Currency
End Type
Function F_BuscaProd(Optional busca As String = "N", _
                 Optional sIni As String = "") As Produto
    
   ' On Error Resume Next
    frmSelect.txtbxExpr.Text = sIni
    frmSelectProd.Show 1
    
    If frmSelectProd.tag = "" Then
        Unload frmSelectProd
        Exit Function
    Else
        With frmSelectProd.dtPro.Recordset
            
            F_BuscaProd.cod_cat = !cod_cat
            F_BuscaProd.cod_mar = !cod_mar
            F_BuscaProd.cod_pro = !cod_pro
            F_BuscaProd.nome_cat = !nome_cat
            F_BuscaProd.nome_mar = !nome_mar
            F_BuscaProd.nome_pro = !nome_pro
            F_BuscaProd.preco = !preco_pro
            Unload frmSelectProd
        End With
    End If
    
End Function


Function F_BuscaProdSaida(Optional busca As String = "", Optional qtd As String = "1", _
    Optional Veiculo As String, Optional ignoraEmpresa As Boolean = False, Optional cod_cli As String = "", _
    Optional cod_emp As String = "") As Produto
    
    On Error GoTo trata_Erro
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    Dim iCount As Integer
    Dim leituraBalanca As String
    On Error GoTo trata_Erro
    Dim u As Veiculo
    u = UltimoVeiculoProcurado
    Set F_BuscaProdSaida.adoAtac = New ADODB.Recordset
    
    frmProdutoSaida.Visible = False
    frmProdutoSaida.tag = busca
    frmProdutoSaida.txtVeiculo.tag = Veiculo
    frmProdutoSaida.cod_cli = cod_cli
    frmProdutoSaida.txtbxBuscar.Text = busca
    frmProdutoSaida.dtgrd.tag = tipo_busca 'código de barra
    frmProdutoSaida.txtQtd.Text = qtd
    frmProdutoSaida.txtEmpresa.tag = cod_emp
    frmProdutoSaida.txtEmpresa.Text = cod_emp
    If UltimoVeiculoProcurado.id > 0 Then
        frmProdutoSaida.txtVeiculo.Text = u.descricao
        frmProdutoSaida.txtVeiculo.tag = u.id
        frmProdutoSaida.lblFab.caption = u.fabricante.descricao
        UltimoVeiculoProcurado = u

    End If
    
    If qtd = "" Then
        frmProdutoSaida.pcQtd.Visible = False
    End If
    frmProdutoSaida.ignora_empresa_qtd_produto = ignoraEmpresa
    frmProdutoSaida.Visible = False
    frmProdutoSaida.Show 1
    
    
    If frmProdutoSaida.tag = "" Then
        Unload frmProdutoSaida
        
        Exit Function
    Else
        With frmProdutoSaida.ado
            
            
            F_BuscaProdSaida.cod_cat = !cod_cat
            F_BuscaProdSaida.cod_mar = !cod_mar
            
            F_BuscaProdSaida.cod_pro = !cod_pro
            F_BuscaProdSaida.codbarra_pro = TiraNull(!codbarra_pro)
            F_BuscaProdSaida.ef_pro = !ef
            F_BuscaProdSaida.er_pro = !er
            F_BuscaProdSaida.nome_cat = Format(!nome_cat)
            F_BuscaProdSaida.nome_mar = Format(!nome_mar)
            F_BuscaProdSaida.cr_pro = !cr_pro
            F_BuscaProdSaida.r_pro = !reservado
            F_BuscaProdSaida.nome_pro = !nome_pro
            F_BuscaProdSaida.pv_pro = !pv_pro
            F_BuscaProdSaida.pa_pro = !pa_pro
            F_BuscaProdSaida.codfab_pro = !codfab_pro
            F_BuscaProdSaida.ncm_pro = Format(!ncm_pro, "0#######")
            
            F_BuscaProdSaida.lucro_pro = !lucro_pro
            
            'F_BuscaProdSaida.pq_pro = !pq_pro
            F_BuscaProdSaida.s_pro = !icms_pro
            F_BuscaProdSaida.ecf_icms = !icms_pro
            F_BuscaProdSaida.codbarra_pro = !codbarra_pro
            F_BuscaProdSaida.balanca_pro = !balanca_pro
            F_BuscaProdSaida.qtd = frmProdutoSaida.txtQtd.Text
            F_BuscaProdSaida.cod_emp = !cod_emp
            F_BuscaProdSaida.ean_pro = !ean_pro
            F_BuscaProdSaida.cat.codigo = !cod_cat
            F_BuscaProdSaida.cat.nome = !nome_cat
            F_BuscaProdSaida.cat.descmax = !descmax_cat
            
            F_BuscaProdSaida.balanca_pro = !balanca_pro
            F_BuscaProdSaida.unidade_pro = !unidade_pro
            F_BuscaProdSaida.controlaestoque_pro = !controlaestoque_pro
            F_BuscaProdSaida.multiplo_pro = !multiplo_pro
            
            
            If !balanca_pro = "S" Then
                If F_SoNumero(busca) = busca Then
                    If Len(busca) = 13 Then
                    
            
                        F_BuscaProdSaida.totalbalanca_pro = CCur(Mid(busca, 8, 5)) / 100
                        F_BuscaProdSaida.qtd = Format(F_BuscaProdSaida.totalbalanca_pro / !pv_pro, "###,##0.000")
                    End If
                End If
            End If
            
            
            
            Unload frmProdutoSaida
        End With
    End If
    If F_BuscaProdSaida.qtd = "" Then
        F_BuscaProdSaida.qtd = "1"
    End If
    
    Exit Function
trata_Erro:
    
End Function
Function buscaProduto(cod_pro) As Produto
    On Error GoTo trata_Erro
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    Dim iCount As Integer
    
    strSQL = "select * from " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " left join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " left join " & TAB_CATEGORIAS & " as tc on tp.cod_cat=tc.cod_cat "
    strSQL = strSQL & " where (cod_pro='" & cod_pro & "' or codfab_pro='" & cod_pro & "'"
    strSQL = strSQL & " or ean_pro='" & cod_pro & "')"
    abreObjetoAdo rPesq, strSQL
    
        With rPesq
            If .RecordCount = 0 Then
                    .Close
                    Exit Function
            End If
            
            buscaProduto.cod_cat = !cod_cat
            buscaProduto.cod_mar = !cod_mar
            
            buscaProduto.cod_pro = !cod_pro
            buscaProduto.codbarra_pro = TiraNull(!codbarra_pro)
            buscaProduto.ef_pro = !ef_pro
            buscaProduto.er_pro = !er_pro
            buscaProduto.nome_cat = Format(!nome_cat)
            buscaProduto.nome_mar = Format(!nome_mar)
            buscaProduto.cr_pro = !cr_pro
            buscaProduto.r_pro = !reservado_pro
            buscaProduto.nome_pro = !nome_pro
            buscaProduto.pv_pro = !pv_pro
            buscaProduto.pa_pro = !pa_pro
            buscaProduto.codfab_pro = !codfab_pro
            buscaProduto.ncm_pro = Format(!ncm_pro, "0#######")
            
            buscaProduto.lucro_pro = !lucro_pro
            
            'buscaProduto.pq_pro = !pq_pro
            buscaProduto.s_pro = !icms_pro
            buscaProduto.ecf_icms = !icms_pro
            buscaProduto.codbarra_pro = !codbarra_pro
            buscaProduto.balanca_pro = !balanca_pro
             
            'BuscaProduto.cod_emp = !cod_emp
            buscaProduto.ean_pro = !ean_pro
            buscaProduto.cat.codigo = !cod_cat
            buscaProduto.cat.nome = !nome_cat
            buscaProduto.cat.descmax = !descmax_cat
            buscaProduto.cat.qtdmin = !qtdmin_cat
            
            
            buscaProduto.balanca_pro = !balanca_pro
            buscaProduto.unidade_pro = !unidade_pro
            buscaProduto.controlaestoque_pro = !controlaestoque_pro
            
            
            
            .Close
            
            
                    End With
    
    Exit Function
trata_Erro:
    
End Function


Function F_InfoProd(cod_pro As String) As InfoProd
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select num_not,entrada_not,tf.cod_for,tf.nome_for,ti.* "
    strSQL = strSQL & " from tab_cmp as tc, tab_itcmp as ti, "
    strSQL = strSQL & " " & TAB_FORNECEDORES & " as tf "
    strSQL = strSQL & " where tf.cod_for=tc.cod_for "
    strSQL = strSQL & " and tc.seq_not=ti.seq_not "
    strSQL = strSQL & " and ti.cod_pro='" & cod_pro & "' "
    strSQL = strSQL & " order by entrada_not desc limit 1 "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount > 0 Then

        
            F_InfoProd.data = !entrada_not
            F_InfoProd.cod_For = !cod_For
            F_InfoProd.cod_pro = !cod_pro
            F_InfoProd.nome_For = !nome_For
            F_InfoProd.ult_custo = !custo_ite
            F_InfoProd.qtd = !qtd_ite * !qtd_emb
            F_InfoProd.unit = !unit_ite
            F_InfoProd.nota = !Num_not
        End If
        .Close
    End With
    Exit Function
trata_Erro:
    MostraErro

End Function
Function buscaProduto_atacado(qtd As String, codigo As String) As produtoAtacado

    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select * from produto_atacado where (codigo ='" & codigo & "' "
    strSQL = strSQL & " or barra='" & codigo & "')"
    
  '  strSQL = strSQL & " and mod(" & F_GravaDec(qtd) & ",qtd)=0 order by qtd desc "
   ' strSQL = strSQL & " and qtd=" & F_GravaDec(qtd)
    
    abreObjetoAdo rPesq, strSQL
    
    
    
    
    
    If rPesq.RecordCount > 0 Then
        buscaProduto_atacado.codigo = rPesq!codigo
        buscaProduto_atacado.id = rPesq!id
        buscaProduto_atacado.preco = rPesq!preco
        buscaProduto_atacado.qtd = rPesq!qtd
    End If
    If rPesq.RecordCount > 1 Then
        strSQL = InputBox("Foram encontradas" & rPesq.RecordCount & " embalagens. Informe a quantidade.", "Informe a quantidade", "1")
        If strSQL <> "" Then
            buscaProduto_atacado.qtd = strSQL
        End If
    End If
    
    FechaObjeto rPesq
    Exit Function
trata_Erro:
    MostraErro
    


End Function


Function F_SelectInfoProd(cod_pro As String) As String
    Dim strSQL As String
    
    On Error GoTo trata_Erro
    strSQL = "select num_not,entrada_not,tf.cod_for,tf.nome_for,ti.* "
    strSQL = strSQL & " from tab_cmp as tc, tab_itcmp as ti, "
    strSQL = strSQL & " " & TAB_FORNECEDORES & " as tf "
    strSQL = strSQL & " where tf.cod_for=tc.cod_for "
    strSQL = strSQL & " and tc.seq_not=ti.seq_not "
    strSQL = strSQL & " and ti.cod_pro='" & cod_pro & "' "
    strSQL = strSQL & " order by entrada_not desc  "
    
    F_SelectInfoProd = strSQL
    Exit Function
trata_Erro:
    MostraErro

End Function



Sub GravaProdutoAtacado(codigo As String, qtd As String, preco As String, barra As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into produto_atacado(codigo,qtd,preco,barra)values("
    strSQL = strSQL & "'" & codigo & "',"
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & F_GravaDec(preco) & ","
    strSQL = strSQL & "'" & barra & "')"
    conn.Execute strSQL
    
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Sub AlteraProdutoAtacado(id As String, qtd As String, preco As String, barra As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "update produto_atacado set  "
    strSQL = strSQL & "barra='" & barra & "', "
    strSQL = strSQL & "qtd=" & F_GravaDec(qtd) & ","
    strSQL = strSQL & "preco=" & F_GravaDec(preco) & " "
    strSQL = strSQL & " where id=" & id
    conn.Execute strSQL
    
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Sub ExcluiProdutoAtacado(id As String)
    Dim strSQL As String
    strSQL = "delete from produto_atacado where id=" & id
    conn.Execute strSQL
End Sub
Sub alteraNCM(cod_pro As String, ncm As String)
    Dim strSQL As String
    strSQL = "update " & TAB_PRODUTOS & " set ncm_pro='" & ncm & "'"
    strSQL = strSQL & " where cod_pro='" & cod_pro & "'"
    conn.Execute strSQL
End Sub
Function F_BuscaSubSec(descricao As String) As subsecao
    frmEdiSubSec.tag = "x"
    'frmEdiSubSec.campo1 = "id"
    'frmEdiSubSec.campo2 = "nome"
    'frmEdiSubSec.campo3 = "id_secao"
    'frmEdiSubSec.tabela = "subsecao"
    frmEdiSubSec.txbxExpr.Text = descricao
    frmEdiSubSec.Show 1
    If frmEdiSubSec.tag <> "" Then
        F_BuscaSubSec.id = frmEdiSubSec.ado!cod_sub
        F_BuscaSubSec.nome = frmEdiSubSec.ado!nome_sub
        F_BuscaSubSec.id_secao = frmEdiSubSec.ado!cod_sec
        F_BuscaSubSec.nome_secao = frmEdiSubSec.ado!nome_sec
    End If
    Unload frmEdiSubSec
End Function
Sub AbrirSimilares(cod_pro As String)
    frmCadSim.tag = cod_pro
    
    frmCadSim.Show 1
End Sub
Function gravaSimimilar(cod_pro As String, cod_sim As String, tipo_sim As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tab_similares(cod_pro,cod_sim,tipo_sim)values("
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & "'" & cod_sim & "',"
    strSQL = strSQL & "'" & tipo_sim & "')"
    conn.Execute strSQL
    gravaSimimilar = True
    Exit Function
trata_Erro:
    
End Function
Function ExcluirSimilar(cod_pro As String, cod_sim As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "delete from tab_similares "
    strSQL = strSQL & " where (cod_pro='" & cod_pro & "' "
    strSQL = strSQL & " and cod_sim='" & cod_sim & "' or cod_pro='" & cod_sim & "' and cod_sim='" & cod_pro & "' )"
    
    conn.Execute strSQL
    ExcluirSimilar = True
    Exit Function
trata_Erro:
    
End Function

Sub AtualizaProdutos(cod_pro As String, campo As String, valor As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update " & TAB_PRODUTOS & " set " & campo & "='" & valor & "' "
    strSQL = strSQL & " where cod_pro='" & cod_pro & "' "
    conn.Execute strSQL
    Exit Sub
trata_Erro:
    MostraErro
    
End Sub
Function buscaSimilares(cod_pro As String) As String
    Dim strSQL As String
    strSQL = "select distinct tp.nome_pro,tp.codfab_pro,tp.cod_pro,tp.pv_pro,tp.cr_pro,tm.nome_mar,tc.nome_cat,ts.nome_sub "
    strSQL = strSQL & " from tab_similares as sim, " & TAB_PRODUTOS & " as tp left join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " left join tab_categorias as tc on tp.cod_cat=tc.cod_Cat "
    strSQL = strSQL & " left join tab_subcat as ts on tp.cod_sub=ts.cod_sub "
    strSQL = strSQL & " where (sim.cod_pro='" & cod_pro & "' or sim.cod_sim='" & cod_pro & "') "
    
    strSQL = strSQL & " and (tp.cod_pro=sim.cod_pro or tp.cod_pro=sim.cod_sim) "
    
    buscaSimilares = strSQL
    
    
End Function
Function BuscaEstoqueSim(cod_pro As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select sum(tp.er_pro) as total from " & TAB_PRODUTOS & " as tp left join tab_similares as ts on tp.cod_pro=ts.cod_pro "
    
    strSQL = strSQL & " where (ts.cod_pro='" & cod_pro & "'  or ts.cod_sim='" & cod_pro & "')"
    strSQL = strSQL & " and ts.cod_pro<>'" & cod_pro & "' "
    abreObjetoAdo rPesq, strSQL
    
    BuscaEstoqueSim = FormatVal(rPesq!total, "###,##0")
    rPesq.Close
    
    
End Function

Sub gravaProdutoProcuradoSemEstoque(cod_pro As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = " insert into produto_procurado(codigo,data,cod_usu)values("
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "')"
    conn.Execute strSQL
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Public Function GravaProduto(codbarra_pro As String, nome_pro As String, unidade_pro As String, cod_cat As String, _
    cod_mar As String, cod_sub As String, subsecao As String, codfab_pro As String, lucro_pro As String, _
    endereco_pro As String, cod_For As String, estmin_pro As String, detalhe_pro As String, _
    local_Pro As String, status_pro As String, ncm_pro As String, ef_pro As String, er_pro As String, _
    cr_pro As String, cf_pro As String, pv_pro As String, pa_pro As String, balanca_pro As String, _
    cod_imp As String, visivel_pro As String, cod_gen As String, cod_ori As String, ean As String, icms_pro As String, _
    cst_icms As String, cst_revenda As String, cfop_pro As String, cfop_revenda As String, CST_IPI As String, cst_pis As String, cst_cofins As String, _
    cEst As String, controlaestoque_pro As String, Optional codigo_produto As String = "", Optional qConv As String = "1", Optional multiplo_pro As String = "1") As String
    Dim strSQL As String
    Dim cod_pro As String
    
    If subsecao = "" Then subsecao = "0"
    
    On Error GoTo trata_Erro
    If codigo_produto = "" Then
    
        sCod = MaxKeySQL(TAB_PRODUTOS, "cod_pro", "", "0000000#")
    Else
        'é migração
        sCod = codigo_produto
        
    End If
    strSQL = " insert into " & TAB_PRODUTOS & "(cod_pro,codbarra_pro,nome_pro,unidade_pro,cod_cat,cod_mar,cod_sub,"
    strSQL = strSQL & " subsecao,codfab_pro,lucro_pro,endereco_pro,cod_for,estmin_pro,detalhe_pro,local_Pro,"
    strSQL = strSQL & "status_pro,ncm_pro,ef_pro,er_pro,cr_pro,cf_pro,pv_pro,pa_pro,balanca_pro,visivel_pro,"
    strSQL = strSQL & " cod_gen,cod_ori,ean_pro,icms_pro,cst_icms,cst_revenda,cfop_pro,cfop_revenda,cst_ipi,cst_pis,cst_cofins,"
    strSQL = strSQL & " cest_pro,controlaestoque_pro,qconv,multiplo_pro)values("
    strSQL = strSQL & "'" & sCod & "',"

    strSQL = strSQL & "'" & codbarra_pro & "',"

    strSQL = strSQL & "'" & nome_pro & "',"

    strSQL = strSQL & "'" & unidade_pro & "',"

    strSQL = strSQL & "'" & cod_cat & "',"

    strSQL = strSQL & "'" & cod_mar & "',"

    strSQL = strSQL & "'" & cod_sub & "',"

    strSQL = strSQL & "'" & subsecao & "',"
    strSQL = strSQL & "'" & codfab_pro & "',"

    strSQL = strSQL & F_GravaDec(lucro_pro) & ", "
    
    strSQL = strSQL & "'" & endereco_pro & "',"

    strSQL = strSQL & "'" & cod_For & "',"

    strSQL = strSQL & F_GravaDec(estmin_pro) & ", "
    strSQL = strSQL & "'" & detalhe_pro & "',"

    strSQL = strSQL & "'" & local_Pro & "',"

    strSQL = strSQL & "'" & status_pro & "',"

    strSQL = strSQL & "'" & ncm_pro & "',"

    strSQL = strSQL & F_GravaDec(ef_pro) & ","
    strSQL = strSQL & F_GravaDec(er_pro) & ", "
    strSQL = strSQL & F_GravaDec(cr_pro) & ", "
    strSQL = strSQL & F_GravaDec(cf_pro) & ", "
    strSQL = strSQL & F_GravaDec(pv_pro) & ", "
    strSQL = strSQL & F_GravaDec(pa_pro) & ", "
    strSQL = strSQL & "'" & balanca_pro & "',"

    

    strSQL = strSQL & "'" & visivel_pro & "',"

    strSQL = strSQL & "'" & cod_gen & "',"
    strSQL = strSQL & "'" & cod_ori & "',"
    strSQL = strSQL & "'" & ean & "',"

    strSQL = strSQL & "'" & icms_pro & "',"

    strSQL = strSQL & "'" & cst_icms & "',"
    strSQL = strSQL & "'" & cst_revenda & "',"
    strSQL = strSQL & "'" & cfop_pro & "',"
    strSQL = strSQL & "'" & cfop_revenda & "',"
    strSQL = strSQL & "'" & CST_IPI & "',"

    strSQL = strSQL & "'" & cst_pis & "',"

    strSQL = strSQL & "'" & cst_cofins & "',"
    strSQL = strSQL & "'" & cEst & "',"
    strSQL = strSQL & "'" & controlaestoque_pro & "',"
    strSQL = strSQL & F_GravaDec(qConv) & ", "
    strSQL = strSQL & F_GravaDec(multiplo_pro) & " "
    strSQL = strSQL & ")"
    conn.Execute strSQL

    GravaProduto = sCod
    Exit Function
trata_Erro:
P_MostraErro
End Function

Function AlteraProduto(cod_pro As String, codbarra_pro As String, nome_pro As String, unidade_pro As String, _
    cod_cat As String, cod_mar As String, cod_sub As String, subsecao As String, _
        codfab_pro As String, lucro_pro As String, endereco_pro As String, _
        cod_For As String, estmin_pro As String, detalhe_pro As String, local_Pro As String, _
        status_pro As String, ncm_pro As String, ef_pro As String, er_pro As String, _
        cr_pro As String, cf_pro As String, pv_pro As String, pa_pro As String, balanca_pro As String, _
        cod_imp As String, visivel_pro As String, cod_gen As String, cod_ori As String, ean As String, icms_pro As String, _
        cst_icms As String, cst_revenda As String, cfop_pro As String, cfop_revenda As String, CST_IPI As String, cst_pis As String, cst_cofins As String, cEst As String, _
        controlaestoque_pro As String, qtdembcompra_pro As String, Optional qConv As String = "1", Optional multiplo_pro As String = "1") As Boolean
    Dim strSQL As String
    
    
    strSQL = " update " & TAB_PRODUTOS & " set "
    
    strSQL = strSQL & "codbarra_pro='" & codbarra_pro & "',"
    strSQL = strSQL & "nome_pro='" & nome_pro & "',"
    strSQL = strSQL & "unidade_pro='" & unidade_pro & "',"
    strSQL = strSQL & "cod_cat='" & cod_cat & "',"
    strSQL = strSQL & "cod_mar='" & cod_mar & "',"
    strSQL = strSQL & "cod_sub='" & cod_sub & "',"
    If subsecao <> "" Then
        strSQL = strSQL & "subsecao=" & subsecao & ", "
    End If
    strSQL = strSQL & "codfab_pro='" & codfab_pro & "',"
    strSQL = strSQL & "lucro_pro=" & str(lucro_pro) & ","
    strSQL = strSQL & "endereco_pro='" & endereco_pro & "',"
    strSQL = strSQL & "cod_for='" & cod_For & "',"
    strSQL = strSQL & "estmin_pro=" & str(estmin_pro) & ","
    strSQL = strSQL & "detalhe_pro='" & detalhe_pro & "',"
    strSQL = strSQL & "local_Pro='" & local_Pro & "',"
    strSQL = strSQL & "status_pro='" & status_pro & "',"
    strSQL = strSQL & "ncm_pro='" & ncm_pro & "',"
    
    'strSQL = strSQL & "ef_pro=" & str(ef_pro) & ","
    
    'strSQL = strSQL & "er_pro=" & str(er_pro) & ", "
    
    
    strSQL = strSQL & "cr_pro=" & str(cr_pro) & ","
    strSQL = strSQL & "qconv=" & str(qConv) & ","
    strSQL = strSQL & "cf_pro=" & str(cf_pro) & ","
    strSQL = strSQL & "pv_pro=" & str(pv_pro) & ","
    strSQL = strSQL & "pa_pro=" & str(pa_pro) & ","
    strSQL = strSQL & "balanca_pro='" & balanca_pro & "',"
    'strSQL = strSQL & "cod_imp='" & cod_imp & "',"
    strSQL = strSQL & "visivel_pro='" & visivel_pro & "',"
    strSQL = strSQL & "cod_gen='" & cod_gen & "',"
    strSQL = strSQL & "cod_ori='" & cod_ori & "',"
    strSQL = strSQL & "ean_pro='" & ean & "',"
    strSQL = strSQL & "cest_pro='" & cEst & "',"
    strSQL = strSQL & "icms_pro='" & icms_pro & "',"
    strSQL = strSQL & "cst_icms='" & cst_icms & "',"
    strSQL = strSQL & "cst_ipi='" & CST_IPI & "',"
    strSQL = strSQL & "cst_pis='" & cst_pis & "',"
    strSQL = strSQL & "cst_revenda='" & cst_revenda & "',"
    strSQL = strSQL & "cfop_pro='" & cfop_pro & "',"
    strSQL = strSQL & "cfop_revenda='" & cfop_revenda & "',"
    strSQL = strSQL & "controlaestoque_pro='" & controlaestoque_pro & "',"
    strSQL = strSQL & "cst_cofins='" & cst_cofins & "',"
    strSQL = strSQL & "qtdembcompra_pro=" & F_GravaDec(qtdembcompra_pro) & ", "
    strSQL = strSQL & "multiplo_pro=" & F_GravaDec(multiplo_pro)
    strSQL = strSQL & " where cod_pro='" & cod_pro & "' "
    conn.Execute strSQL
    AlteraProduto = True
    Exit Function
trata_Erro:

End Function
Function AlteraOpcoesNfe(cod_pro As String, cod_gen As String, cod_ori As String, icms_pro As String, _
        cst_icms As String, CST_IPI As String, cst_pis As String, cst_cofins As String, cEst As String) As Boolean
    Dim strSQL As String
    
    
    
    
    strSQL = " update " & TAB_PRODUTOS & " set "
    
    If cod_gen <> "" Then
        strSQL = strSQL & "cod_gen='" & cod_gen & "',"
    End If
    If cod_ori <> "" Then
        strSQL = strSQL & "cod_ori='" & cod_ori & "',"
    End If
    If cEst <> "" Then
        strSQL = strSQL & "cest_pro='" & cEst & "',"
    End If
    If icms_pro <> "" Then
        strSQL = strSQL & "icms_pro='" & icms_pro & "',"
    End If
    If cst_icms <> "" Then
        strSQL = strSQL & "cst_icms='" & cst_icms & "',"
    End If
    If CST_IPI <> "" Then
        If CST_IPI = "-1" Then
            CST_IPI = ""
        End If
        
        strSQL = strSQL & "cst_ipi='" & CST_IPI & "',"
    
        
    End If
    If cst_pis <> "" Then
        If cst_pis = "-1" Then
            cst_pis = ""
        End If
        strSQL = strSQL & "cst_pis='" & cst_pis & "',"
    End If
    If cst_cofins <> "" Then
        If cst_cofins = "-1" Then
            cst_cofins = ""
        End If
        strSQL = strSQL & "cst_cofins='" & cst_cofins & "',"
    End If
    strSQL = strSQL & " nome_pro=nome_pro "
    strSQL = strSQL & " where cod_pro='" & cod_pro & "' "
    conn.Execute strSQL
    AlteraOpcoesNfe = True
    Exit Function
trata_Erro:

End Function

Sub AlteraUltimoCusto(cod_pro As String, uc_pro As String)
    strSQL = "update " & TAB_PRODUTOS & " set uc_pro=" & F_GravaDec(uc_pro)
    strSQL = strSQL & " where cod_pro='" & cod_pro & "'"
    conn.Execute strSQL
End Sub
Function F_AlteraEstRes(cod_pro As String, qtd As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro:
    
    
    
    strSQL = "update " & TAB_PRODUTOS & " set r_pro=r_pro+" & F_GravaDec(qtd)
    strSQL = strSQL & "  where cod_pro='" & cod_pro & "' "
    conn.Execute strSQL
    
    
    F_AlteraEstRes = True
    
    Exit Function
trata_Erro:
P_MostraErro
End Function
Function F_AlteraER_Pro(cod_pro As String, er_pro As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro:
    
    
    
    strSQL = "update " & TAB_PRODUTOS & " set er_pro=" & F_GravaDec(er_pro)
    strSQL = strSQL & "  where cod_pro='" & cod_pro & "' "
    conn.Execute strSQL
    
    
    F_AlteraER_Pro = True
    
    Exit Function
trata_Erro:
P_MostraErro
End Function
Function F_AlteraEF_Pro(cod_pro As String, ef_pro As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro:
    
    
    
    strSQL = "update " & TAB_PRODUTOS & " set ef_pro=" & F_GravaDec(ef_pro)
    strSQL = strSQL & "  where cod_pro='" & cod_pro & "' "
    conn.Execute strSQL
    
    
    F_AlteraEF_Pro = True
    
    Exit Function
trata_Erro:
P_MostraErro
End Function



Function F_BuscaUltimoPrecoCompra(cod_pro As String) As String
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select unit_ite from tab_cmp as tc, tab_itcmp as ti"
    strSQL = strSQL & " where tc.seq_not=ti.seq_not "
    strSQL = strSQL & " and ti.cod_pro='" & cod_pro & "' "
    strSQL = strSQL & " order by emissao_not desc limit 1 "
    abreObjetoAdo rPesq, strSQL
    F_BuscaUltimoPrecoCompra = rPesq!unit_ite
    
    FechaObjeto rPesq
    
    
Exit Function
trata_Erro:
F_BuscaUltimoPrecoCompra = "0"
End Function

Sub P_HistProd(cod_pro As String)
    frmHistProd.tag = cod_pro
    frmHistProd.txtProduto.Text = F_BuscaUmCampo("nome_pro", TAB_PRODUTOS, "cod_pro='" & cod_pro & "'")
    frmHistProd.txtProduto.tag = cod_pro
    frmHistProd.Show 1

End Sub
Function AdicionaEstoque(cod_emp As String, cod_pro As String, er As String, ef As String, cfop_revenda As String, cfop_consumidor As String, _
    cst_icms_revenda As String, cst_icms_consumidor As String, Optional cst_pis As String = "49", Optional cst_cofins As String = "49", _
    Optional CST_IPI As String = "", Optional pICMS As String = "0", Optional predIcms As String = "0", _
    Optional pPIS As String = "0", Optional pCOFINS As String = "0") As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into estoque(cod_emp,cod_pro,er,ef,cfop_revenda,cfop_consumidor,cst_icms_revenda,cst_icms_consumidor,"
    strSQL = strSQL & "cst_pis,cst_cofins,cst_ipi,picms,predicms,ppis,pcofins)values("
    strSQL = strSQL & "'" & cod_emp & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(er) & ","
    strSQL = strSQL & F_GravaDec(ef) & ","
    strSQL = strSQL & "'" & cfop_revenda & "',"
    strSQL = strSQL & "'" & cfop_consumidor & "',"
    strSQL = strSQL & "'" & cst_icms_revenda & "',"
    strSQL = strSQL & "'" & cst_icms_consumidor & "',"
    strSQL = strSQL & "'" & cst_pis & "',"
    strSQL = strSQL & "'" & cst_cofins & "',"
    strSQL = strSQL & "'" & CST_IPI & "',"
    strSQL = strSQL & F_GravaDec(pICMS) & ","
    strSQL = strSQL & F_GravaDec(pPIS) & ","
    strSQL = strSQL & F_GravaDec(pCOFINS) & ","
    strSQL = strSQL & F_GravaDec(predIcms) & ")"
    conn.Execute strSQL
    AdicionaEstoque = True
trata_Erro:
    
End Function
Function AlteraEstoque(cod_emp As String, cod_pro As String, er As String, ef As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update estoque set "
    
    strSQL = strSQL & "er=" & F_GravaDec(er) & ","
    strSQL = strSQL & "ef=" & F_GravaDec(ef)
    strSQL = strSQL & " where cod_emp='" & cod_emp & "' and "
    strSQL = strSQL & " cod_pro='" & cod_pro & "'"
    conn.Execute strSQL
    AlteraEstoque = True
trata_Erro:
    
End Function
Function AlteraDadosEstoque(cod_emp As String, cod_pro As String, cfop_revenda As String, cfop_consumidor As String, _
    cst_icms_revenda As String, cst_icms_consumidor As String, Optional cst_pis As String = "49", Optional cst_cofins As String = "49", _
    Optional CST_IPI As String = "", Optional pICMS As String = "0", Optional predIcms As String = "0", _
    Optional pPIS As String = "0", Optional pCOFINS As String = "0") As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update estoque set "
    
    strSQL = strSQL & "cfop_revenda='" & cfop_revenda & "',"
    strSQL = strSQL & "cfop_consumidor='" & cfop_consumidor & "',"
    strSQL = strSQL & "cst_icms_revenda='" & cst_icms_revenda & "',"
    strSQL = strSQL & "cst_icms_consumidor='" & cst_icms_consumidor & "',"
    strSQL = strSQL & "cst_pis='" & cst_pis & "',"
    strSQL = strSQL & "cst_cofins='" & cst_cofins & "',"
    strSQL = strSQL & "cst_ipi='" & CST_IPI & "',"
    strSQL = strSQL & "picms=" & F_GravaDec(pICMS) & ","
    strSQL = strSQL & "ppis=" & F_GravaDec(pPIS) & ","
    strSQL = strSQL & "pcofins=" & F_GravaDec(pCOFINS) & ","
    strSQL = strSQL & "predicms=" & F_GravaDec(predIcms)
    strSQL = strSQL & " where cod_emp='" & cod_emp & "' and "
    strSQL = strSQL & " cod_pro='" & cod_pro & "'"
    conn.Execute strSQL
    AlteraDadosEstoque = True
trata_Erro:
    
End Function



Function AlteraEstoqueProduto(cod_pro As String, er As String, ef As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update estoque set "
    
    strSQL = strSQL & "er=" & F_GravaDec(er) & ","
    strSQL = strSQL & "ef=" & F_GravaDec(ef)
    strSQL = strSQL & " where cod_pro='" & cod_pro & "'"
    conn.Execute strSQL
    AlteraEstoqueProduto = True
trata_Erro:
    
End Function

Function BaixaEstoque(cod_emp As String, cod_pro As String, er As String, ef As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update estoque set "
    
    strSQL = strSQL & "er=er-" & F_GravaDec(er) & ","
    strSQL = strSQL & "ef=ef-" & F_GravaDec(ef)
    strSQL = strSQL & " where cod_emp='" & cod_emp & "' and "
    strSQL = strSQL & " cod_pro='" & cod_pro & "',"
    conn.Execute strSQL
    BaixaEstoque = True
trata_Erro:
    
End Function



Sub AbreProdutoProcurado(cod_pro As String)
    frmProdutosProcuradosSemEstoque.tag = cod_pro
    frmProdutosProcuradosSemEstoque.Show 1
End Sub
Sub P_BuscaEstoqueReservado(cod_pro As String)
    frmRelEstoqueReservado.caption = cod_pro
    frmRelEstoqueReservado.txtProd.Text = F_BuscaUmCampo("nome_pro", TAB_PRODUTOS, "cod_pro='" & cod_pro & "'")
    frmRelEstoqueReservado.txtProd.tag = cod_pro
    frmRelEstoqueReservado.tag = cod_pro
    frmRelEstoqueReservado.Show 1
    
End Sub
Sub alteraEAN(cod_pro As String, ean As String)
    Dim strSQL As String
    strSQL = "update " & TAB_PRODUTOS & " set ean_pro='" & ean & "'"
    strSQL = strSQL & " where cod_pro='" & cod_pro & "'"
    conn.Execute strSQL
End Sub

Function ExcluirProduto(cod_pro As String) As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim bDeleta As Boolean
    On Error GoTo trata_Erro
    
    
    
    If Not F_senha Then Exit Function
    If strExclVenda_usu <> "S" Then
        MsgBox "Você não tem permissão!", vbExclamation
        Exit Function
    End If
    
    
    With ado
        
        
        
        strSQL = "select s.*,i.* from sped_nf as s inner join sped_nf_itens  as i on s.id=i.id_nf"
        strSQL = strSQL & " where i.cod_pro='" & cod_pro & "' "
        abreObjetoAdo rPesq, strSQL
        
        If rPesq.RecordCount > 0 Then
            MsgBox "Há uma compra para este produto. SEQ: " & rPesq!id, vbCritical
            Exit Function
        End If
        strSQL = "select * from nfe_item "
        strSQL = strSQL & " where cprod='" & cod_pro & "' "
        abreObjetoAdo rPesq, strSQL
        
        If rPesq.RecordCount > 0 Then
            MsgBox "Há uma compra para este produto. SEQ: " & rPesq!id, vbCritical
            Exit Function
        End If
        
        
        
        
        strSQL = "delete from " & TAB_PRODUTOS & " where cod_pro = '" & cod_pro & "'"
        conn.Execute strSQL
        strSQL = "delete from estoque where cod_pro = '" & cod_pro & "'"
        conn.Execute strSQL
        
        
    End With
        
    ExcluirProduto = True
        
Exit Function
trata_Erro:
MostraErro

End Function

Function GravaProdutoConversao(cod_pro As String, cod_mar As String, referencia As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into produto_conversao(cod_pro,cod_mar,referencia) values ("
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & "'" & cod_mar & "',"
    strSQL = strSQL & "'" & referencia & "')"
    conn.Execute strSQL
    GravaProdutoConversao = True
trata_Erro:
End Function
Function BuscaProdutoConversao(cod_pro As String, cod_mar As String) As ProdutoConversao
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from produto_conversao where cod_pro='" & cod_pro & "' "
    strSQL = strSQL & " and cod_mar='" & cod_mar & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            BuscaProdutoConversao.cod_mar = !cod_mar
            BuscaProdutoConversao.cod_pro = !cod_pro
            BuscaProdutoConversao.id = !id
            BuscaProdutoConversao.referencia = !referencia
        End If
        .Close
    End With
    
trata_Erro:
End Function


Function AlteraProdutoConversao(id As String, referencia As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update produto_conversao set referencia='" & referencia & "' "
    strSQL = strSQL & " where id='" & id & "' "
    conn.Execute strSQL
    AlteraProdutoConversao = True
trata_Erro:
End Function
Function ExcluiProdutoConversao(id As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "delete from produto_conversao  "
    strSQL = strSQL & " where id='" & id & "' "
    conn.Execute strSQL
    ExcluiProdutoConversao = True
trata_Erro:
End Function
Function SelectProdutoConversao(cod_pro As String, Optional cod_mar As String, Optional referencia As String) As String
    Dim strSQL As String
    strSQL = "select tp.*,tm.nome_mar from produto_conversao as tp inner join " & TAB_MARCAS & " as tm "
    strSQL = strSQL & " on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " where 1 > 0 "
    If cod_pro <> "" Then
        strSQL = strSQL & " and cod_pro='" & cod_pro & "' "
    End If
    If cod_mar <> "" Then
        strSQL = strSQL & " and tc.cod_mar='" & cod_mar & "' "
    End If
    strSQL = strSQL & " and referencia like '%" & referencia & "%' "
    
    strSQL = strSQL & " order by nome_mar,referencia  "
    SelectProdutoConversao = strSQL
End Function
Sub P_AltPreco(cod_pro As String, novo_preco As Currency, antigo As Currency, Optional bAlteraPreco As Boolean = False)
    Dim strSQL As String
    On Error GoTo trata_Erro
    If novo_preco = antigo Then
        Exit Sub
    End If
    
    If bAlteraPreco Then
        strSQL = "update " & TAB_PRODUTOS & " set pv_pro=" & F_GravaDec(novo_preco)
        strSQL = strSQL & " where cod_pro='" & cod_pro & "'"
        conn.Execute strSQL
    End If
    
    strSQL = "insert into tab_altpreco(cod_pro,antigo_pro,novo_pro,data_alt,cod_usu)values("
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(antigo) & ","
    strSQL = strSQL & F_GravaDec(novo_preco) & ","
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "')"
    conn.Execute strSQL
    Exit Sub
trata_Erro:
End Sub

Function RemoveProdutoCombustivel(cod_pro As String) As Boolean
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "delete from produtos_comb where cod_pro= '" & cod_pro & "',"
    
    conn.Execute strSQL
    RemoveProdutoCombustivel = True
    Exit Function
trata_Erro:

End Function
Function GravaProdutoCombustivel(cod_pro As String, cProdAnp As String, descAnp As String) As Boolean
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "replace into produtos_comb(cod_pro,cProdAnp,DescAnp)values('" & cod_pro & "',"
    strSQL = strSQL & "'" & cProdAnp & "',"
    strSQL = strSQL & "'" & descAnp & "') "
    conn.Execute strSQL
    GravaProdutoCombustivel = True
    Exit Function
trata_Erro:

End Function
