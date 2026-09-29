Attribute VB_Name = "Produtos"
Public Type Estoque
    
    seq_pro As String
    ef_pro As Currency
    er_pro As Currency
    venc_pro As String
    lote_pro As String

End Type

Public Enum tipoMovimento
    EntradaMovimento = 1
    saidaMovimento = 2
End Enum
Public Enum tipoMovEst
    venda_ = 1
    compra = 2
    Devolucao = 3
    contagem = 4
End Enum
Type Produto1
    cod_pro As String
    ref_pro As String
    codbar_pro As String
    nome_pro As String
    complemento_pro As String
    pv_pro As Currency
    pp_pro As Currency
    prom_pro As String
    prin As principio.principio
    sub As subsecAO.subsec
    est(100) As Estoque
    cr_pro As Currency
    cf_pro As Currency
    cm_pro As Currency
    
    s_pro As String
    lucro_pro As Currency
    
    icms As Currency
    ecf_icms As String
    fracionado_pro As String
    pfab_pro As Currency
    unidade_pro As String
    
    ean_pro As String
    ncm_pro As String
    
End Type

Type controle
    cod_con As String
    nome_con As String
    descricao_con As String
End Type


Type ProdutoSaida
    cod_pro As String
    codbar_pro As String
    nome_pro As String
    ref_pro As String
    pv_pro As Currency
    pp_pro As Currency
    prom_pro As String
    prin As principio.principio
    sub As subsecAO.subsec
    est As Estoque
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
    ean_pro As String
    ncm_pro As String
End Type
Type produto
    cod_pro As String
    codbar_pro As String
    codbarra_pro As String
    nome_pro As String
    ref_pro As String
    pv_pro As Currency
    pp_pro As Currency
    prom_pro As String
    prin As principio.principio
    sub As subsecAO.subsec
    est As Estoque
    cr_pro As Currency
    cf_pro As Currency
    cm_pro As Currency
    cod_sub As String
    s_pro As String
    lucro_pro As Currency
    endereco_pro As String
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
    ean_pro As String
    cb_pro As Currency
    ncm_pro As String
    cfop_pro As String
    unidade_pro As String
    cst_pis As String
    cst_cofins As String
    CST_IPI As String
    cst_icms As String
    cod_ori As String
    cfop_revenda As String
    cst_revenda As String
    id_pro As String
End Type

Type ProdutoConsig
    seq_con As String
    pro As produtos.ProdutoSaida
    data_con As String
    qtd_con As Integer
    cod_Rep As String
    cod_usu As String
    reman_con As Currency
    unit_con As Currency
End Type

Function F_BuscaProdConsign(cod_Rep As String, Optional busca As String) As ProdutoConsig
    On Error Resume Next
    frmProdConsig.tag = busca
    frmProdConsig.lblNomeRep.tag = cod_Rep
    frmProdConsig.lblNomeRep.caption = cod_Rep
    frmProdConsig.Show 1
    If frmProdConsig.tag = "" Then
    Else
        With frmProdConsig.adoprod
            F_BuscaProdConsign.cod_Rep = !cod_Rep
            F_BuscaProdConsign.cod_usu = !cod_usu
            F_BuscaProdConsign.data_con = !data_con
            F_BuscaProdConsign.seq_con = !seq_con
            F_BuscaProdConsign.qtd_con = !qtd_con
            'F_BuscaProdConsign.reman_con = !reman_con
            F_BuscaProdConsign.reman_con = !Restante 'já tirando o reservado do que foi lançado em aberto
            
            If CCur(!Restante) = 0 And CCur(!reman_con) > 0 Then
                MsgBox "Este produto está lançado  em acertos que estão em aberto restando apenas  " & !Restante, vbExclamation
            End If
            
            F_BuscaProdConsign.unit_con = !unit_con
        
                F_BuscaProdConsign.pro.cod_pro = !cod_pro
                F_BuscaProdConsign.pro.fracionado_pro = !fracionado_pro
                F_BuscaProdConsign.pro.pa_pro = !pa_pro
                F_BuscaProdConsign.pro.lucro_pro = !lucro_pro
                F_BuscaProdConsign.pro.s_pro = !st_pro
                F_BuscaProdConsign.pro.cm_pro = !cm_pro
                F_BuscaProdConsign.pro.codbar_pro = !codbarra_pro
                F_BuscaProdConsign.pro.cr_pro = !cr_pro
                F_BuscaProdConsign.pro.pq_pro = !pq_pro
                               
                
                
                
                F_BuscaProdConsign.pro.sub.cod_sub = !cod_sub
                F_BuscaProdConsign.pro.sub.nome_sub = !nome_sub
                F_BuscaProdConsign.pro.sub.descmax_sub = !descmax_sub
                F_BuscaProdConsign.pro.sub.comiss_sub = !comiss_sub
                F_BuscaProdConsign.pro.sub.sec.nome_sec = !nome_sec
                F_BuscaProdConsign.pro.pv_pro = !pv_pro
                F_BuscaProdConsign.pro.cr_pro = !cr_pro
                F_BuscaProdConsign.pro.ecf_icms = F_IcmsProd(!cod_pro, 1)
                F_BuscaProdConsign.pro.nome_pro = !nome_pro
                F_BuscaProdConsign.pro.ref_pro = !ref_pro
                F_BuscaProdConsign.pro.complemento_pro = !complemento_pro
                F_BuscaProdConsign.pro.pfab_pro = !pfab_pro
                .Close
        
        End With
    End If
    Unload frmProdConsig
    
    
    
    
End Function

Function F_BuscaProdSaida(Optional busca As String = "", Optional tipo_busca As String = "N") As produto
    On Error Resume Next
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    Dim iCount As Integer
    On Error GoTo trata_Erro
    Dim adoData As New ADODB.Recordset
    
      
    
    
    
    
    frmProdutoSaida.tag = busca
    
    frmProdutoSaida.dtgrd.tag = tipo_busca 'código de barra
    frmProdutoSaida.Show 1
    
    If frmProdutoSaida.tag = "" Then
        Unload frmProdutoSaida
        
        Exit Function
    Else
        Set adoData = frmProdutoSaida.dtgrd.DataSource
        With adoData
            
            F_BuscaProdSaida.cod_pro = !cod_pro
                
                F_BuscaProdSaida.fracionado_pro = !fracionado_pro
                F_BuscaProdSaida.pa_pro = !pa_pro
                F_BuscaProdSaida.lucro_pro = !lucro_pro
                F_BuscaProdSaida.s_pro = !st_pro
                F_BuscaProdSaida.cm_pro = !cm_pro
                If Not IsNull(!codbarra_pro) Then
                    F_BuscaProdSaida.codbar_pro = !codbarra_pro
                End If
                F_BuscaProdSaida.cr_pro = !cr_pro
                
                               
                
                
                
                F_BuscaProdSaida.sub.cod_sub = !cod_sub
                F_BuscaProdSaida.sub.nome_sub = !nome_sub
                F_BuscaProdSaida.sub.descmax_sub = !descmax_sub
                F_BuscaProdSaida.sub.comiss_sub = !comiss_sub
                F_BuscaProdSaida.sub.sec.nome_sec = !nome_sec
                F_BuscaProdSaida.pv_pro = !pv_pro
                F_BuscaProdSaida.cr_pro = !cr_pro
                If !icms_pro = "FF" Or !icms_pro = "II" Then
                    F_BuscaProdSaida.ecf_icms = 0
                Else
                    F_BuscaProdSaida.ecf_icms = CCur(!icms_pro)
                End If
                
                F_BuscaProdSaida.nome_pro = !nome_pro
                F_BuscaProdSaida.ref_pro = !ref_pro
                If Not IsNull(!complemento_pro) Then
                    F_BuscaProdSaida.complemento_pro = !complemento_pro
                End If
                F_BuscaProdSaida.pfab_pro = !pfab_pro
                F_BuscaProdSaida.est.er_pro = !er_pro
                F_BuscaProdSaida.est.ef_pro = !ef_pro
                F_BuscaProdSaida.pq_pro = !pq_pro
                
            Unload frmProdutoSaida
            
        End With
    End If
    Exit Function
trata_Erro:
    P_MostraErro "Busca de Produtos"
End Function

Function F_GravaProduto(nome_pro As String, cod_sub As String, ref_pro As String, cf_pro As String, cr_pro As String, _
    pv_pro As String, pa_pro As String, unidade_pro As String, _
    pfab_pro As String, cod_fab As String, endereco_pro As String, st_pro As String, icms_pro As String, _
    lucro_pro As String, cb_pro As String, pq_pro As String, principal_pro As String, francionado_pro As String, ncm As String, ean_pro As String, codbarra_pro As String) As String
    Dim sCod As String
    Dim strSQL As String
    
    
    sCod = MaxKeySQL("tab_produtos", "cod_pro", "", "0000000#")

    If Trim(ref_pro) = "" Then
        ref_pro = sCod
    End If
    
'    If sCod = "" Then
'        sCod = "0000001"
'    Else
'        sCod = Format(Int(Left(sCod, 7)) + 1, "000000#")
'    End If
'    sCod = F_Ean8(sCod)
'conn.BeginTrans
    strSQL = "insert into tab_produtos(Cod_pro,nome_pro,cod_sub,"
    strSQL = strSQL & " ref_pro,cr_pro,cf_pro,pv_pro,pa_pro,unidade_pro,principal_pro,"
    strSQL = strSQL & " pfab_pro,cod_fab,endereco_pro,st_pro,icms_pro,lucro_pro,cb_pro,pq_pro,fracionado_pro,ncm_pro,"
    strSQL = strSQL & " dataalt_pro,ean_pro,codbarra_pro)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_pro & "',"
    strSQL = strSQL & "'" & cod_sub & "',"
    strSQL = strSQL & "'" & ref_pro & "',"
    strSQL = strSQL & F_GravaDec(cr_pro) & ","
    strSQL = strSQL & F_GravaDec(cf_pro) & ","
    strSQL = strSQL & F_GravaDec(pv_pro) & ","
    strSQL = strSQL & F_GravaDec(pa_pro) & ","
    strSQL = strSQL & "'" & unidade_pro & "', "
    strSQL = strSQL & "'" & principal_pro & "', "
    strSQL = strSQL & F_GravaDec(pfab_pro) & ", "
    strSQL = strSQL & "'" & cod_fab & "', "
    strSQL = strSQL & "'" & endereco_pro & "', "
    strSQL = strSQL & "'" & st_pro & "', "
    strSQL = strSQL & "'" & icms_pro & "', "
    strSQL = strSQL & F_GravaDec(lucro_pro) & ", "
    strSQL = strSQL & F_GravaDec(cb_pro) & ", "
    strSQL = strSQL & F_GravaDec(pq_pro) & ", "
    strSQL = strSQL & "'" & francionado_pro & "', "
    strSQL = strSQL & "'" & ncm & "', "
    strSQL = strSQL & F_GravaData(Date) & ", "
    strSQL = strSQL & "'" & ean_pro & "', "
    strSQL = strSQL & "'" & codbarra_pro & "') "
    Conn.Execute strSQL
    
    
    If produtos.F_AddCodBar(sCod, sCod) Then
    End If
 '   conn.CommitTrans
    
    
    F_GravaProduto = sCod
    
    Exit Function
trata_Erro:

End Function
Sub P_GravaProdutoComCodigo(cod_pro As String, nome_pro As String, cod_sub As String, ref_pro As String, cf_pro As String, cr_pro As String, _
    pv_pro As String, pa_pro As String, unidade_pro As String, _
    pfab_pro As String, cod_fab As String, endereco_pro As String, icms_pro As String, _
    lucro_pro As String, cb_pro As String, pq_pro As String, st_pro As String)
    Dim sCod As String
    Dim strSQL As String
    sCod = cod_pro
    strSQL = "insert into tab_produtos(Cod_pro,nome_pro,cod_sub,"
    strSQL = strSQL & " ref_pro,cr_pro,cf_pro,pv_pro,pa_pro,unidade_pro,"
    strSQL = strSQL & " pfab_pro,cod_fab,endereco_pro,icms_pro,lucro_pro,cb_pro,pq_pro,st_pro)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_pro & "',"
    strSQL = strSQL & "'" & cod_sub & "',"
    strSQL = strSQL & "'" & ref_pro & "',"
    strSQL = strSQL & F_GravaDec(cr_pro) & ","
    strSQL = strSQL & F_GravaDec(cf_pro) & ","
    strSQL = strSQL & F_GravaDec(pv_pro) & ","
    strSQL = strSQL & F_GravaDec(pa_pro) & ","
    strSQL = strSQL & "'" & unidade_pro & "', "
    strSQL = strSQL & F_GravaDec(pfab_pro) & ", "
    strSQL = strSQL & "'" & cod_fab & "', "
    strSQL = strSQL & "'" & endereco_pro & "', "
    strSQL = strSQL & "'" & icms_pro & "', "
    strSQL = strSQL & F_GravaDec(lucro_pro) & ", "
    strSQL = strSQL & F_GravaDec(cb_pro) & ", "
    strSQL = strSQL & F_GravaDec(pq_pro) & ","
    strSQL = strSQL & "'" & st_pro & "') "
    Conn.Execute strSQL
    
    If produtos.F_AddCodBar(sCod, sCod) Then
    End If
    
    
    
    
    
    Exit Sub
trata_Erro:

End Sub


Function F_MigraProduto(cod_pro As String, nome_pro As String, cod_sub As String, ref_pro As String, cf_pro As String, cr_pro As String, _
    pv_pro As String, pa_pro As String, unidade_pro As String, _
    pfab_pro As String, cod_fab As String, endereco_pro As String, icms_pro As String, lucro_pro As String) As String
    Dim sCod As String
    Dim strSQL As String
    sCod = Format(cod_pro, "0000000#")
    strSQL = "insert into tab_produtos(Cod_pro,nome_pro,cod_sub,"
    strSQL = strSQL & " ref_pro,cr_pro,cf_pro,pv_pro,pa_pro,unidade_pro,"
    strSQL = strSQL & " pfab_pro,cod_fab,endereco_pro,icms_pro,lucro_pro)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_pro & "',"
    strSQL = strSQL & "'" & cod_sub & "',"
    strSQL = strSQL & "'" & ref_pro & "',"
    strSQL = strSQL & F_GravaDec(cr_pro) & ","
    strSQL = strSQL & F_GravaDec(cf_pro) & ","
    strSQL = strSQL & F_GravaDec(pv_pro) & ","
    strSQL = strSQL & F_GravaDec(pa_pro) & ","
    strSQL = strSQL & "'" & unidade_pro & "', "
    strSQL = strSQL & F_GravaDec(pfab_pro) & ", "
    strSQL = strSQL & "'" & cod_fab & "', "
    strSQL = strSQL & "'" & endereco_pro & "', "
    strSQL = strSQL & "'" & icms_pro & "', "
    strSQL = strSQL & F_GravaDec(lucro_pro) & ") "
    Conn.Execute strSQL
    
    If produtos.F_AddCodBar(sCod, sCod) Then
    End If
    
    
    
    F_MigraProduto = sCod
    
    Exit Function
trata_Erro:

End Function





Function F_EditaProduto(cod_pro As String, nome_pro As String, cod_sub As String, ref_pro As String, cr_pro As String, _
    pv_pro As String, pa_pro As String, unidade_pro As String, _
    pfab_pro As String, cod_fab As String, endereco_pro As String, st_pro As String, _
    icms_pro, cb_pro As String, pq_pro As String, fracionado_pro As String, principal_pro As String, ncm_pro As String, _
    cod_gen As String, cod_ori As String, ean As String, _
        cst_icms As String, CST_IPI As String, cst_pis As String, cst_cofins As String, cEst As String, ean_pro As String, codbarra_pro As String) As Boolean
    
    On Error GoTo trata_Erro:
    Dim i As Integer
    If ref_pro = "" Then
        ref_pro = cod_pro
    End If
    strSQL = "update tab_produtos set nome_pro='" & nome_pro & "',"
    
    strSQL = strSQL & "cod_sub='" & cod_sub & "',"
    
    strSQL = strSQL & "ref_pro='" & ref_pro & "',"
    strSQL = strSQL & "cr_pro=" & F_GravaDec(cr_pro) & ","
    strSQL = strSQL & "cb_pro=" & F_GravaDec(cb_pro) & ","
    strSQL = strSQL & "pq_pro=" & F_GravaDec(pq_pro) & ","
    strSQL = strSQL & "pv_pro=" & F_GravaDec(pv_pro) & ", "
    strSQL = strSQL & "pa_pro=" & F_GravaDec(pa_pro) & ", "
    
    strSQL = strSQL & "fracionado_pro='" & fracionado_pro & "', "
    
    strSQL = strSQL & "unidade_pro='" & unidade_pro & "' ,"
    strSQL = strSQL & "principal_pro='" & principal_pro & "' ,"
    strSQL = strSQL & "pfab_pro=" & F_GravaDec(pfab_pro) & ","
    strSQL = strSQL & "cod_fab='" & cod_fab & "', "
    
    
    strSQL = strSQL & "endereco_pro='" & endereco_pro & "', "
    strSQL = strSQL & "lista_pro='" & lista_pro & "', "
    strSQL = strSQL & "st_pro='" & st_pro & "', "
    strSQL = strSQL & "ncm_pro='" & ncm_pro & "', "
    
    strSQL = strSQL & "cod_gen='" & cod_gen & "',"
    strSQL = strSQL & "cod_ori='" & cod_ori & "',"
    strSQL = strSQL & "ean_pro='" & ean & "',"
    strSQL = strSQL & "cest_pro='" & cEst & "',"
    strSQL = strSQL & "icms_pro='" & icms_pro & "',"
    strSQL = strSQL & "cst_icms='" & cst_icms & "',"
    strSQL = strSQL & "cst_ipi='" & CST_IPI & "',"
    strSQL = strSQL & "cst_pis='" & cst_pis & "',"

    strSQL = strSQL & "cst_cofins='" & cst_cofins & "',"
    strSQL = strSQL & "ean_pro='" & ean_pro & "',"
    strSQL = strSQL & "codbarra_pro='" & codbarra_pro & "',"
    
    strSQL = strSQL & "icms_pro='" & icms_pro & "' "
    strSQL = strSQL & " where cod_pro='" & cod_pro & "' "
    Conn.Execute strSQL, i
    If i > 0 Then
        strSQL = "update " & TAB_PRODUTOS & " set dataalt_pro=" & F_GravaData(Date)
        strSQL = strSQL & " where cod_pro='" & cod_pro & "'  "
        Conn.Execute strSQL
    End If
    
    F_EditaProduto = True
    Exit Function
trata_Erro:
    P_MostraErro
    F_EditaProduto = False
End Function


Function F_AddCodBar(cod_pro As String, cod_bar As String) As Boolean
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    If F_BuscaUmCampo("cod_pro", "tab_codbarras", "cod_pro='" & cod_pro & "' and cod_bar='" & cod_bar & "'") <> "" Then
'        MsgBox "Já existe um item com este código de barras para este produto ", vbExclamation
'        F_AddCodBar = False
        Exit Function
    End If
    
    
    strSQL = "insert into tab_codbarras(cod_pro,cod_bar)values("
    
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & "'" & cod_bar & "')"
    
    Conn.Execute strSQL
    F_AddCodBar = True
    Exit Function
trata_Erro:
    P_MostraErro
    F_AddCodBar = False
End Function

Function F_EditaEstoque(cod_pro As String, venc As String, lote As String, seq_pro As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "update tab_estoque set " '(cod_pro,lote_pro,venc_pro,er_pro,ef_pro,ce_pro)values("
    strSQL = strSQL & "lote_pro='" & lote & "',"
    strSQL = strSQL & "venc_pro=" & F_GravaData(venc) & " "
    strSQL = strSQL & " where cod_pro='" & cod_pro & "' "
    strSQL = strSQL & " and seq_pro=" & seq_pro
    Conn.Execute strSQL
    F_EditaEstoque = True
    Exit Function
trata_Erro:
    P_MostraErro
    F_EditaEstoque = False
End Function
Function F_BuscaProdCabecalho(expr As String) As produtos.produto
    frmSelect.sSelect = "select cod_pro as cod,nome_pro as nome, pv_pro as Preco, cr_pro as Custo,   unidade_pro as Un from tab_produtos "
    frmSelect.cCriterio.Add " where codbarra_pro = '" & expr & "'"
    frmSelect.cCriterio.Add " where cod_pro = '" & expr & "'"
    frmSelect.cCriterio.Add " where nome_pro like '" & expr & "%'"
    frmSelect.sOrd = " order by nome_pro "
    frmSelect.txtbxExpr.text = expr
    
    frmSelect.Show 1
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        
        F_BuscaProdCabecalho.cod_pro = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.col = 1
        
        F_BuscaProdCabecalho.nome_pro = frmSelect.msflxgrd
        
        frmSelect.msflxgrd.col = 2
        F_BuscaProdCabecalho.pv_pro = frmSelect.msflxgrd
        
        frmSelect.msflxgrd.col = 3
        F_BuscaProdCabecalho.cr_pro = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 4
        
        F_BuscaProdCabecalho.unidade_pro = frmSelect.msflxgrd
       
        
    End If
    
    
    
    Unload frmSelect
    
End Function

Function F_RetornaProduto(cod_pro) As produtos.produto
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = produtos.F_SelectProduto("tp.cod_pro='" & cod_pro & "'")
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            F_RetornaProduto.cod_pro = !cod_pro
            F_RetornaProduto.lucro_pro = !lucro_pro
                F_RetornaProduto.s_pro = !st_pro
                F_RetornaProduto.cm_pro = !cm_pro
                F_RetornaProduto.codbar_pro = !codbarra_pro
                F_RetornaProduto.cr_pro = !cr_pro
                F_RetornaProduto.est.ef_pro = !ef_pro
                F_RetornaProduto.est.er_pro = !er_pro
                F_RetornaProduto.est.lote_pro = !lote_pro
                F_RetornaProduto.est.venc_pro = !venc_pro
                F_RetornaProduto.est.seq_pro = !seq_pro
                F_RetornaProduto.prin.cod_pri = !cod_pri
                F_RetornaProduto.prin.nome_pri = !nome_pri
                F_RetornaProduto.sub.cod_sub = !cod_sub
                F_RetornaProduto.sub.nome_sub = !nome_sub
                F_RetornaProduto.sub.sec.nome_sec = !nome_sec
                F_RetornaProduto.pv_pro = !pv_pro
                F_RetornaProduto.cr_pro = !cr_pro
                F_RetornaProduto.ecf_icms = F_IcmsProd(!cod_pro, 1)
                F_RetornaProduto.nome_pro = !nome_pro
                F_RetornaProduto.ref_pro = !ref_pro
        End If
        .Close
    End With
    
End Function
Function F_SelectProduto(Optional sWh As String = "") As String
    Dim strSQL As String
    strSQL = "select distinct ts.cod_sec, ts.nome_sec, tu.nome_sub,tf.nome_fab, "
    strSQL = strSQL & "ti.cod_imp,pis.descricao as pis, icms.descricao as icms, cofins.descricao as cofins, "
    strSQL = strSQL & "ipi.descricao as ipi, ti.desc_imp,tp.*, "
    strSQL = strSQL & "gen.descricao as genero,"
    strSQL = strSQL & "tp.*,tp.er_pro as estoque  "
    strSQL = strSQL & " from tab_produtos as tp left join tab_fabricantes as tf "
    strSQL = strSQL & " on tp.cod_fab=tf.cod_fab "
    strSQL = strSQL & " left join tab_subsec as tu "
    strSQL = strSQL & " on tp.cod_sub=tu.cod_sub "
    strSQL = strSQL & " left join tab_secoes as ts "
    strSQL = strSQL & " on tu.cod_sec=ts.cod_sec "

    strSQL = strSQL & " left join sped_genero as gen "
    strSQL = strSQL & " on tp.cod_gen=gen.codigo "
    strSQL = strSQL & " left join " & TAB_IMPOSTOS & " as ti on tp.cod_imp=ti.cod_imp"
    strSQL = strSQL & " left join sped_cst_pis as pis on tp.cst_pis=pis.codigo"
    strSQL = strSQL & " left join sped_cst_cofins as cofins on tp.cst_cofins=cofins.codigo"
    strSQL = strSQL & " left join sped_cst_ipi as ipi on tp.cst_ipi=ipi.codigo"
    strSQL = strSQL & " left join nfe_tipo_icms as icms on tp.cst_icms=icms.codigo"
    
    If sWh <> "" Then
        strSQL = strSQL & " where " & sWh
    Else
        strSQL = strSQL & " where 1 > 0 "
    End If
    
    
    
        
        
        F_SelectProduto = strSQL
End Function
Function F_BuscaProdControle(expr As String) As controle
    frmSelect.sSelect = "select cod_con as cod,nome_con as nome, descricao_con as descrição from tab_controleS "
    frmSelect.cCriterio.Add " where nome_con like '" & expr & "%'"
    frmSelect.sOrd = " order by nome_con "
    frmSelect.txtbxExpr.text = expr
    frmSelect.Show 1
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        
        F_BuscaProdControle.cod_con = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.col = 1
        
        F_BuscaProdControle.nome_con = frmSelect.msflxgrd
        
        frmSelect.msflxgrd.col = 2
        
        F_BuscaProdControle.descricao_con = frmSelect.msflxgrd
        
    End If
    Unload frmSelect
    
End Function

Function F_RetornaNomeControle(cod_con As String) As String
    F_RetornaNomeControle = F_BuscaUmCampo("nome_con", "tab_controles", "cod_con='" & cod_con & "'")
End Function
Function F_BuscaEstoqueProd(cod_pro As String) As Estoque
    frmSelect.sSelect = "select seq_pro as Seq,lote_pro as lote, venc_pro as venc,er_pro as ER, ef_pro as ef from tab_estoque "
    frmSelect.cCriterio.Add " where cod_pro = '" & cod_pro & "'"
    frmSelect.sOrd = " order by seq_pro "
    frmSelect.txtbxExpr.text = expr
    frmSelect.Show 1
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        
        F_BuscaEstoqueProd.seq_pro = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.col = 1
        
        F_BuscaEstoqueProd.lote_pro = frmSelect.msflxgrd
        
        frmSelect.msflxgrd.col = 2
        
        F_BuscaEstoqueProd.venc_pro = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 3
        
        F_BuscaEstoqueProd.er_pro = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 4
        
        F_BuscaEstoqueProd.ef_pro = frmSelect.msflxgrd
    End If
    Unload frmSelect
End Function
Function F_SuspPro(cod_pro As String, motivo_pro As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tab_suspensao(cod_pro,data_pro,cod_usu,motivo_sus)values("
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & motivo_pro & "') "
    Conn.Execute strSQL
    
    F_SuspPro = True
    Exit Function
trata_Erro:
    
End Function
Sub P_BuscaPrecos(preco, nome_pro As String, msf As MSFlexGrid)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    With msf
        .FixedCols = 0
        .rows = 1
        .Cols = 2
        .col = 0: .text = "Plano"
        .col = 1: .text = "Preço"
        
        .ColWidth(0) = 1800
    End With
    msf.rows = 1
    strSQL = "select * from tab_tipoven "
    abreObjetoAdo rPesq, strSQL
    msf.ColAlignment(0) = 3
    msf.ColAlignment(1) = 3
    With rPesq
        Do While Not .EOF
        
            strSQL = !descricao_tip & Chr(9) & FormatVal(CCur(preco) * !fator_tip / !parcelas_tip)
            msf.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    ZebraGrid msf, 1, &HB4DFE0
End Sub

Function F_AlteraEstRes(cod_pro As String, qtd As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro:
    
    
    
    strSQL = "update tab_produtos set r_pro=r_pro+" & F_GravaDec(qtd)
    strSQL = strSQL & "  where cod_pro='" & cod_pro & "' "
    Conn.Execute strSQL
    
    
    F_AlteraEstRes = True
    
    Exit Function
trata_Erro:
P_MostraErro
End Function
Function F_AlteraEstoquePro(cod_pro As String, qtd As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro:
    strSQL = "update tab_produtos set er_pro=er_pro+" & F_GravaDec(qtd)
    strSQL = strSQL & "  where cod_pro='" & cod_pro & "' "
    Conn.Execute strSQL
    
    
    F_AlteraEstoquePro = True
    Exit Function
trata_Erro:
End Function
Sub P_deletaQueima(cod_pro As String)
    Dim strSQL As String
    strSQL = "delete from tab_queima where cod_pro='" & cod_pro & "' "
    Conn.Execute strSQL
End Sub
Sub P_GravaQuei(cod_pro As String, pq_pro As String)
    Dim strSQL As String
    P_deletaQueima cod_pro
    strSQL = "insert into tab_queima(cod_pro,pq_pro)values("
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(pq_pro) & ") "
    Conn.Execute strSQL
End Sub
Sub P_GravaSimilar(cod_pro As String, cod_pro2 As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sCod As String
    Dim sCod2 As String
    sCod = F_BuscaUmCampo("seq_sim", "tab_similares", "cod_pro='" & cod_pro & "'")
    
    sCod2 = F_BuscaUmCampo("seq_sim", "tab_similares", "cod_pro='" & cod_pro2 & "'")
    
    If sCod = "" And sCod2 = "" Then
        sCod = MaxKeySQL("tab_similares", "seq_sim", "", "")
        strSQL = "insert into tab_similares(seq_sim,cod_pro)values("
        strSQL = strSQL & sCod & ","
        strSQL = strSQL & "'" & cod_pro & "')"
        Conn.Execute strSQL
        strSQL = "insert into tab_similares(seq_sim,cod_pro)values("
        strSQL = strSQL & sCod & ","
        strSQL = strSQL & "'" & cod_pro2 & "')"
        Conn.Execute strSQL
        
    ElseIf sCod <> "" Then
        strSQL = "insert into tab_similares(seq_sim,cod_pro)values("
        strSQL = strSQL & sCod & ","
        strSQL = strSQL & "'" & cod_pro2 & "')"
        Conn.Execute strSQL
    Else
    
    End If
    
    
End Sub
Sub P_GeraListaProdutosMigracao()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_produtos order by cod_pro "
    abreObjetoAdo rPesq, strSQL
        
    
    
End Sub
Sub P_ProdutosPorRepresentante(Optional cod_pro As String = "", Optional cod_Rep As String = "")
    frmProdutosPorRepresentante.txtbxRep.text = cod_Rep
    frmProdutosPorRepresentante.txtbxRep.tag = cod_Rep
    frmProdutosPorRepresentante.txtbxProd.text = cod_pro
    frmProdutosPorRepresentante.txtbxProd.tag = cod_pro
    frmProdutosPorRepresentante.tag = "B"
    frmProdutosPorRepresentante.Show 0
    frmProdutosPorRepresentante.SetFocus
    
End Sub
Function F_SelectTodosProdutos(Optional expr As String, Optional cod_sec As String = "", Optional cod_sub As String = "") As String
    Dim strSQL As String
    strSQL = "select distinct ts.nome_sub, tse.nome_sec,"
    strSQL = strSQL & " tp.*,er_pro+r_pro as estoque,((tp.pv_pro/tp.cr_pro-1)*100)-" & F_GravaDec(F_TaxasEntradas) & "  as liq, "
    strSQL = strSQL & " ((tp.pv_pro/tp.cr_pro-1)*100) as lb, "
    strSQL = strSQL & " (((tp.pv_pro/tp.cr_pro-1)*100)-" & F_GravaDec(F_TaxasEntradas) & ") * pv_pro/100  as lr, "
    strSQL = strSQL & " ((tp.pv_pro/tp.cr_pro-1)*100) * pv_pro/100  as bruto "
    strSQL = strSQL & " from tab_produtos as tp,tab_subsec as ts, tab_secoes as tse,tab_codbarras as tb "
    
    strSQL = strSQL & " where tp.cod_pro=tb.cod_pro "
    strSQL = strSQL & " and ts.cod_sec=tse.cod_sec "
    strSQL = strSQL & " and tp.cod_sub=ts.cod_sub "
    If expr <> "" Then
        strSQL = strSQL & " and (nome_pro like '%"
        strSQL = strSQL & expr & "%'  "
        
        If F_SoNumero(expr) <> "" Then
            strSQL = strSQL & " or ref_pro =" & expr
        End If
        strSQL = strSQL & " or codbarra_pro like '" & expr & "%' "
        strSQL = strSQL & " or endereco_pro like '%" & expr & "%' "
        strSQL = strSQL & " or tp.cod_pro like '" & expr & "%'"
        strSQL = strSQL & " or tb.cod_bar='" & expr & "') "
    End If
    If cod_sec <> "" Then
        strSQL = strSQL & " and tse.cod_sec='" & cod_sec & "' "
    End If
    If cod_sub <> "" Then
        strSQL = strSQL & " and ts.cod_sub='" & cod_sub & "' "
    End If
    F_SelectTodosProdutos = strSQL
End Function
Sub P_GeraArquivoTransfProdutos(sCaminho As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sPtc As String
    sPtc = ChaveSQL("tab_produtos", "protocolo_pro", "000000#")
    If sCaminho = "" Then
        mdiFrm.cmndlg.FileName = "P" & sPtc & ".txt"
        mdiFrm.cmndlg.filter = ".txt"
        mdiFrm.cmndlg.ShowSave
        
        
        sCaminho = mdiFrm.cmndlg.FileName
        
    End If
    If sCaminho = "" Then
        sCaminho = App.path & "\P" & seq_mov
        MsgBox "O arquivo será salvo em " & sCaminho
    Else
        
    End If
    strSQL = F_SelectProduto()
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Produtos não encontrados!", vbExclamation
            .Close
            Exit Sub
        End If
        
        Open sCaminho For Output As #1
        Print #1, "0001" 'produtos
        Print #1, "1000" & F_SoNumero(strCNPJ_Emp)
        Print #1, "1001" & sPtc
        Print #1, "1002" & Format(Date, "DD/MM/YYYY")
        Print #1, "1003" & Now
        Print #1, "1004" & strCod_Usu
        Print #1, "1005" & strNome_Usu
        Do While Not .EOF
            Print #1, "2" & Format(.AbsolutePosition, "00#") & F_EsticaTexto(!cod_pro, 8);
            'Print #1, F_EsticaTexto(!seq_ite, 3);
            'Print #1, F_EsticaTexto(!qtd_ite, 8);
            Print #1, F_EsticaTexto(!nome_pro, 50);
            Print #1, F_EsticaTexto(!cod_sub, 3);
            Print #1, F_EsticaTexto(!nome_sub, 20);
            Print #1, F_EsticaTexto(!cod_sec, 3);
            Print #1, F_EsticaTexto(!nome_sec, 20);
            Print #1, F_EsticaTexto(!cr_pro, 8);
            Print #1, F_EsticaTexto(!cf_pro, 8);
            Print #1, F_EsticaTexto(!cb_pro, 8); 'custo bruto
            Print #1, F_EsticaTexto(!pv_pro, 8); 'preço venda
            Print #1, F_EsticaTexto(!lucro_pro, 8);
            Print #1, F_EsticaTexto(!st_pro, 1);
            Print #1, F_EsticaTexto(!icms_pro, 2);
            Print #1, F_EsticaTexto(!unidade_pro, 3)
            Print #1, F_EsticaTexto(!cod_fab, 5)
            Print #1, F_EsticaTexto(F_BuscaUmCampo("nome_for", "tab_fornecedores", "cod_for='" & !cod_fab & "'"), 50)
            
            .MoveNext
        Loop
        Print #1, "9999"
        .Close
        Close #1
    End With
    MsgBox "Arquivo salvo " & sCaminho, vbInformation
    
End Sub

Function F_CalcLucroAplicado(preco As Currency, custo As Currency)
    On Error Resume Next
    F_CalcLucroAplicado = (preco - custo) / custo * 100
    
End Function
Sub P_GravaHistEst(cod_pro As String, Anterior As String, atual As String, doc As String, tipo As tipoMovEst)
    'tipo :
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tab_histest(cod_pro,anterior_pro,atual_pro,num_doc,tipo_doc)values("
    strSQL = strSQL & "'" & cod_pro & ","
    strSQL = strSQL & F_GravaDec(Anterior) & ","
    strSQL = strSQL & F_GravaDec(atual) & ","
    strSQL = strSQL & "'" & doc & "',"
    strSQL = strSQL & "'" & tipo & "')"
    Conn.Execute strSQL
    
    Exit Sub
trata_Erro:
P_MostraErro "Gravação do Histórico da Movimentação"
End Sub
Sub alteraNCM(cod_pro As String, ncm As String)
    Dim strSQL As String
    strSQL = "update " & TAB_PRODUTOS & " set ncm_pro='" & ncm & "'"
    strSQL = strSQL & " where cod_pro='" & cod_pro & "'"
    Conn.Execute strSQL
End Sub
Sub alteraEAN(cod_pro As String, ean As String)
    Dim strSQL As String
    strSQL = "update " & TAB_PRODUTOS & " set ean_pro='" & ean & "'"
    strSQL = strSQL & " where cod_pro='" & cod_pro & "'"
    Conn.Execute strSQL
End Sub

Sub ImprEtiquetaPromocional(cod_pro As String)
    ImprimeEtiquetaPromocional cod_pro
    
    End Sub
Sub ImprimeEtiquetaPromocional(cod_pro As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim arquivo As String
    Dim style As String
    
    strSQL = "select cod_pro,tp.nome_pro, tp.pv_pro, tp.pq_pro from " & TAB_PRODUTOS & " as tp"
    
    strSQL = strSQL & " where cod_pro='" & cod_pro & "' "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount <= 0 Then
            MsgBox "Produto não encontrado!", vbExclamation
            .Close
            Exit Sub
        End If
        arquivo = html.F_AbreHtml(!cod_pro & "_promo", "Promoção")
        html.P_HtmlAbreBody
        style = "style=" & aspas & " display: block;margin-left: auto; margin-right: auto;width=40%;" & aspas
        
        html.IMG 1, App.path & "\pic\logo_promocao.png", style
  
        'html.P_HtmlBR
        html.Paragrafo 1, !nome_pro, True, "center", "48", "Times", "black", "bold"
        html.Paragrafo 1, "De: " & FormatVal(!pv_pro), True, "center", "38", "Times", "black", "normal", "line-through"
        html.Paragrafo 1, "Por: " & FormatVal(!pq_pro), True, "center", "50", "Times", "red", "bold", "normal"
        html.Paragrafo 1, "Desconto: " & FormatVal((1 - (!pq_pro / !pv_pro)) * 100, "###,##0") & "%", True, "center", "26", "Times", "black", "bold"
        html.Paragrafo 1, "À vista/cartão ", True, "center", "25", "Times", "black", "bold", "underline"
        html.Paragrafo 1, "Código: " & !cod_pro, True, "center", "22", "Times", "black", "normal"
        
 
        style = "style=" & aspas & "display: block;margin-left: auto; margin-right: auto;width=20px;" & aspas
        html.IMG 1, App.path & "\pic\logoprint.jpg", style
        html.P_HtmlFechaBody
        
        html.P_FechaHtml 1, True, arquivo
                
        .Close
    End With
    
    
   

End Sub
Sub GeraArquivoMigracaoProduto(Optional cod_pro = "", Optional validade_dias As Integer = 1)
    
    Dim produtos
    Dim conteudo
    Dim data As String
    If cod_pro <> "" Then
        cod_pro = " where cod_pro in (" & cod_pro & ")"
    End If
    data = DateAdd("D", validade_dias, Date)
    P_Status "processando registros"
    produtos = TabToJson("tab_produtos" & cod_pro, "", "", 0, 0, True, "", False)
    
    P_Status ""
    conteudo = "{" & """data""" & " : " & """" & Now & """" & "," & CrLf
    conteudo = conteudo & """validade""" & " : " & """" & EncriptarToken(data, "REQUINTE") & """" & "," & CrLf
    conteudo = conteudo & """GERADO POR""" & " : " & """" & strNome_Usu & """" & "," & CrLf
    conteudo = conteudo & """tables""" & ": [" & CrLf & "{" & CrLf
    conteudo = conteudo & """table""" & ":" & crl & """tab_produtos""" & "," & CrLf
    conteudo = conteudo & """list_produtos""" & " : " & CrLf & "["
    conteudo = conteudo & produtos & "]}]}"
    If GravaArquivoExternoUTF8("", conteudo, True, "json") Then
    End If
    
    
    
End Sub

Sub importarProdutos(Optional conteudo = "", Optional avisar As Boolean = False)
    
    Dim c As New Collection
    Dim node
    Dim partesJson
    Dim str As String
    Dim p As produto
    Dim val As String
    Dim strSQL As String
    Dim vencimento As String
    Dim pos As Long
    Dim saidaSQL
   ' On Error GoTo trata_Erro
    On Error Resume Next
    saidaSQL = ""
    If conteudo = "" Then
        conteudo = LeArquivoExternoUTF8("")
    
    
    
        vencimento = ExtrairValorObjetoJSON(conteudo, "validade")
        If vencimento = "" Then
            MsgBox "Não foi possível encontrar o vencimento do arquivo de importação", vbExclamation
            Exit Sub
        End If
        vencimento = DesencriptarToken(vencimento, "REQUINTE")
        
        If Date > CDate(vencimento) Then
            MsgBox "Não foi possível importar os produtos. Validade expirada em " & vencimento, vbExclamation
            Exit Sub
        End If
    
        node = JSONUtil.RetornaNodeJson("list_produtos", "", conteudo)
    Else
        node = conteudo
    End If
    partesJson = Split(node, "{")
    Conn.BeginTrans
    Set c = RetornaLinhasJson(node)
    For Each valor In c
        pos = pos + 1
        p.cod_pro = ExtrairValorObjetoJSON(valor, "cod_pro")
        p.id_pro = ExtrairValorObjetoJSON(valor, "id_pro")
        If p.cod_pro = "00000812" Then
'           MsgBox ""
        End If
        val = (ExtrairValorObjetoJSON(valor, "cf_pro", True, False))
        If val <> "" And val <> "null" Then
            p.cf_pro = val
        Else
            val = "0"
        End If
        val = ExtrairValorObjetoJSON(valor, "cm_pro", True, False)
        If val <> "" Then
            p.cm_pro = CCur(val)
        Else
            p.cm_pro = "0"
        End If
        p.codbarra_pro = ExtrairValorObjetoJSON(valor, "gtin_pro")
        p.cod_fab = ExtrairValorObjetoJSON(valor, "cod_fabrica")
        p.codbar_pro = ExtrairValorObjetoJSON(valor, "codbar_pro")
        p.complemento_pro = ExtrairValorObjetoJSON(valor, "complemento_pro")
        cr = ExtrairValorObjetoJSON(valor, "cr_pro", True, False)
        If cr <> "" And cr <> "null" Then
            p.cr_pro = cr
        Else
            p.cr_pro = "0"
        End If
        p.ean_pro = ExtrairValorObjetoJSON(valor, "ean_pro")
        p.ref_pro = p.cod_pro
        val = ExtrairValorObjetoJSON(valor, "fracionado_pro")
        If val = "t" Then
            p.fracionado_pro = "S"
        Else
            p.fracionado_pro = "N"
        End If
                    
        
        val = ExtrairValorObjetoJSON(valor, "margem_lucro", True, IIf(InStr(1, valor, ".") > 0, True, False))
        If val <> "" Then
            p.lucro_pro = CCur(val) * 100
        End If
        p.unidade_pro = ExtrairValorObjetoJSON(valor, "unidade_pro")
        
'        p.pfab_pro = ExtrairValorObjetoJSON(valor, "pfab_pro", True, True)
        val = ExtrairValorObjetoJSON(valor, "pv_pro", True, False)
        p.pv_pro = val
        'p.pa_pro = ExtrairValorObjetoJSON(valor, "pa_pro", True, True)
        val = ExtrairValorObjetoJSON(valor, "pq_pro", True, False)
        If val <> "" Then
            p.pq_pro = CCur(val)
        End If
        'p.endereco_pro = ExtrairValorObjetoJSON(valor, "endereco_pro")
        val = ExtrairValorObjetoJSON(valor, "ef_pro", True, True)
        'p.est.ef_pro = ExtrairValorObjetoJSON(valor, "ef_pro", True, True)
        'p.est.er_pro = ExtrairValorObjetoJSON(valor, "er_pro", True, True)
        p.ncm_pro = ExtrairValorObjetoJSON(valor, "ncm_pro")
        val = ExtrairValorObjetoJSON(valor, "cfop_pro")
        If val <> "" Then
            p.cfop_pro = val
        Else
            p.cfop_pro = "5102"
        End If
        p.cst_cofins = ExtrairValorObjetoJSON(valor, "cst_cofins")
        val = ExtrairValorObjetoJSON(valor, "cst_icms")
        If val = "" Then val = "102"
        
        p.cst_icms = val
        
        p.cst_pis = ExtrairValorObjetoJSON(valor, "cst_pis")
        p.cod_ori = ExtrairValorObjetoJSON(valor, "cod_ori")
        p.cod_sub = Format(ExtrairValorObjetoJSON(valor, "cod_sub"), "00#")
        p.nome_pro = F_AjustaTexto(ExtrairValorObjetoJSON(valor, "nome_pro"))
        
        p.cfop_revenda = ExtrairValorObjetoJSON(valor, "cfop_revenda")
        p.cst_revenda = ExtrairValorObjetoJSON(valor, "cst_revenda")
        
        strSQL = SQLReplaceProdutoSemEstoque(p)
        Conn.Execute strSQL
        saidaSQL = saidaSQL & strSQL & ";" & CrLf
        
        strSQL = SQLUpdateSemEstoque(p)
        Conn.Execute strSQL
        
        gravaAtualizacaoProduto p.cod_pro
        
       ' saidaSQL = saidaSQL & strSQL & ";" & CrLf
        
        If avisar = True Then
            If pos Mod 100 = 0 Then
                P_Status "migrando produto " & p.cod_pro
            End If
        End If
    Next
   
    
    Conn.CommitTrans
   
    GravaArquivoExterno App.path & "\saida_migracao.sql", saidaSQL, False, "sql"
    If avisar Then
        P_Status
    End If
    'MsgBox pos & " produtos importados", vbInformation
    
    Exit Sub
trata_Erro:
    Conn.RollbackTrans
    P_MostraErro
    
End Sub
Sub gravaAtualizacaoProduto(cod_pro As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "replace into atualizacao_produtos(cod_pro,atualizacao)values("
    strSQL = strSQL & "'" & cod_pro & "', Now() )"
    Conn.Execute strSQL
    
    Exit Sub
trata_Erro:
    
    
    
End Sub

Function SQLUpdateSemEstoque(p As produto) As String
    Dim strSQL As String
    strSQL = "update " & TAB_PRODUTOS & " set "
    strSQL = strSQL & "cr_pro=" & F_GravaDec(p.cr_pro) & ","
    strSQL = strSQL & "cf_pro=" & F_GravaDec(p.cf_pro) & ","
    strSQL = strSQL & "pv_pro=" & F_GravaDec(p.pv_pro) & ","
    
    strSQL = strSQL & "pq_Pro=" & F_GravaDec(p.pq_pro) & ","
    strSQL = strSQL & "cm_pro=" & F_GravaDec(p.cm_pro) & ","
    strSQL = strSQL & "ref_pro='" & p.ref_pro & "',"
    strSQL = strSQL & "fracionado_pro='" & p.fracionado_pro & "',"
    strSQL = strSQL & "lucro_pro=" & F_GravaDec(p.lucro_pro) & ","
    If p.cod_fab <> "" Then
        strSQL = strSQL & "cod_fab='" & p.cod_fab & "',"
    End If
    strSQL = strSQL & "unidade_Pro='" & p.unidade_pro & "',"
    
    
    
    'strSQL = strSQL & F_GravaDec(p.cb_pro) & ","
    strSQL = strSQL & "ncm_pro='" & p.ncm_pro & "'"
    If p.cfop_pro <> "" Then
        strSQL = strSQL & ",cfop_pro='" & p.cfop_pro & "' "
    End If
    If p.CST_IPI <> "" Then
        strSQL = strSQL & ",cst_ipi='" & p.CST_IPI & "' "
    End If
    If p.cst_pis <> "" Then
        strSQL = strSQL & ",cst_pis='" & p.cst_pis & "' "
    End If
    If p.cst_cofins <> "" Then
        strSQL = strSQL & ",cst_cofins='" & p.cst_cofins & "' "
    End If
    If p.ean_pro <> "" Then
        strSQL = strSQL & ",ean_pro='" & p.ean_pro & "' "
    End If
    If p.cst_icms <> "" Then
        strSQL = strSQL & ",cst_icms='" & p.cst_icms & "' "
    End If
    If p.cfop_revenda <> "" Then
        strSQL = strSQL & ",cfop_revenda='" & p.cfop_revenda & "' "
    End If
    If p.cst_revenda <> "" Then
        strSQL = strSQL & ",cst_revenda='" & p.cst_revenda & "' "
    End If
    strSQL = strSQL & " where cod_pro='" & p.cod_pro & "' "
    SQLUpdateSemEstoque = strSQL
End Function

Function SQLReplaceProdutoSemEstoque(p As produto) As String
    Dim strSQL As String
    strSQL = "replace into " & TAB_PRODUTOS & "(cod_pro,nome_pro,cod_sub)values("
    strSQL = strSQL & "'" & p.cod_pro & "',"
    strSQL = strSQL & "'" & p.nome_pro & "',"
    strSQL = strSQL & "'" & p.cod_sub & "')"
    
    SQLReplaceProdutoSemEstoque = strSQL
End Function


