Attribute VB_Name = "Financeiro"
Type cheque1
    num_che As String
    cod_bco As String
    cco_che As String
    age_che As String
    nome_che As String
    cnpj_che As String
    Venc_che As String
    Valor_che As Currency
    cpf_che As String
End Type
Type chePag1 '
    num_che As String
    cod_con As String
    Venc_che As String
    Valor_che As Currency
End Type
Type parcela
    vencimento As String
    valor As Currency
End Type
Type localche
    cod_loc As String
    nome_loc As String
End Type
Type ContaCorrente
    cod_con As String
    nome_con As String

End Type

Function F_SelectContasPagar(Optional dataI As String = "", Optional dataF As String = "", _
    Optional num_doc As String = "", Optional seq_doc As String = "", Optional seq_itd As String = "", _
    Optional cod_for As String = "", Optional cod_pla As String = "", Optional historico_doc As String = "", _
    Optional emaberto As Boolean = True, Optional sOrderBy As String = "order by venc_itd,nome_for") As String
    Dim strSQL As String
    
    strSQL = "select tf.nome_for,tf.cnpj_for,td.*,ti.*,tp.nome_pla,concat(td.parc_doc,'/',ti.parc_itd) as parc,"
    strSQL = strSQL & " tf.nome_for as nome,tf.cnpj_for as cnpj "
    strSQL = strSQL & " from tab_doc1 as td, tab_planocontas as tp, tab_fornecedores as tf, "
    strSQL = strSQL & " tab_itdoc1 as ti "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.cod_pla=tp.cod_pla "
    strSQL = strSQL & " and td.cod_ter=tf.cod_for "
    strSQL = strSQL & " and td.tipo_doc='P' "
    If dataI <> "" Then
        strSQL = strSQL & " and venc_itd>=" & F_GravaData(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and venc_itd<=" & F_GravaData(dataF)
    End If
    If num_doc <> "" Then
        strSQL = strSQL & " and num_doc='" & num_doc & "' "
    End If
    If seq_doc <> "" Then
        strSQL = strSQL & " and td.seq_doc='" & seq_doc & "' "
    End If
    If seq_itd <> "" Then
        strSQL = strSQL & " and ti.seq_itd='" & seq_itd & "' "
    End If
    If cod_for <> "" Then
        strSQL = strSQL & " and tf.cod_for='" & cod_for & "' "
    End If
    If cod_pla <> "" Then
        strSQL = strSQL & " and td.cod_pla='" & cod_pla & "' "
    End If
    If historico_doc <> "" Then
        strSQL = strSQL & " and td.historico_doc='" & historico_doc & "' "
    End If
    If emaberto Then
        strSQL = strSQL & " and ti.reman_itd>0"
    End If
    If sOrderBy <> "" Then
        strSQL = strSQL & sOrderBy
    End If
    F_SelectContasPagar = strSQL
End Function

Function F_SelectContasReceber(Optional dataI As String = "", Optional dataF As String = "", _
    Optional num_doc As String = "", Optional seq_doc As String = "", Optional seq_itd As String = "", _
    Optional cod_cli As String = "", Optional cod_pla As String = "", Optional historico_doc As String = "", _
    Optional emaberto As Boolean = True, Optional sOrderBy As String = "order by venc_itd,nome_cli") As String
    Dim strSQL As String
    
    strSQL = "select tc.nome_cli,tc.cpf_cli,td.*,ti.*,tp.nome_pla,concat(td.parc_doc,'/',ti.parc_itd) as parc,"
    strSQL = strSQL & " tc.nome_cli as nome,tc.cpf_cli as cpf  "
    strSQL = strSQL & " from tab_doc1 as td, tab_planocontas as tp, tab_clientes as tc, "
    strSQL = strSQL & " tab_itdoc1 as ti "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.cod_pla=tp.cod_pla "
    strSQL = strSQL & " and td.cod_ter=tc.cod_cli "
    strSQL = strSQL & " and td.tipo_doc='R' "
    If dataI <> "" Then
        strSQL = strSQL & " and venc_itd>=" & F_GravaData(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and venc_itd<=" & F_GravaData(dataF)
    End If
    If num_doc <> "" Then
        strSQL = strSQL & " and num_doc='" & num_doc & "' "
    End If
    If seq_doc <> "" Then
        strSQL = strSQL & " and td.seq_doc='" & seq_doc & "' "
    End If
    If seq_itd <> "" Then
        strSQL = strSQL & " and ti.seq_itd='" & seq_itd & "' "
    End If
    If cod_cli <> "" Then
        strSQL = strSQL & " and tc.cod_cli='" & cod_cli & "' "
    End If
    If cod_pla <> "" Then
        strSQL = strSQL & " and td.cod_pla='" & cod_pla & "' "
    End If
    If historico_doc <> "" Then
        strSQL = strSQL & " and td.historico_doc='" & historico_doc & "' "
    End If
    If emaberto Then
        strSQL = strSQL & " and ti.reman_itd>0"
    End If
    If sOrderBy <> "" Then
        strSQL = strSQL & sOrderBy
    End If
    F_SelectContasReceber = strSQL
End Function


Function f_BuscaLocalChe(expr As String) As localche
    frmSelect.sSelect = "select cod_loc as cod,nome_loc as nome from tab_localche "
    frmSelect.cCriterio.Add " where nome_loc like '" & expr & "%'"
    frmSelect.sOrd = " order by nome_loc "
    frmSelect.txtbxExpr.Text = expr
    frmSelect.Show 1
    If frmSelect.Tag <> "" Then
        frmSelect.msflxgrd.col = 0
        
        f_BuscaLocalChe.cod_loc = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.col = 1
        
        f_BuscaLocalChe.nome_loc = frmSelect.msflxgrd
        
        
        
    End If
    Unload frmSelect
End Function
Function F_BuscaContaCorrente(expr As String) As ContaCorrente
    frmSelect.sSelect = "select cod_con as cod,nome_con as nome from tab_contas "
    frmSelect.cCriterio.Add " where nome_con like '" & expr & "%'"
    frmSelect.sOrd = " order by nome_con "
    frmSelect.txtbxExpr.Text = expr
    frmSelect.Show 1
    If frmSelect.Tag <> "" Then
        frmSelect.msflxgrd.col = 0
        
        F_BuscaContaCorrente.cod_con = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.col = 1
        
        F_BuscaContaCorrente.nome_con = frmSelect.msflxgrd
        
        
        
    End If
    Unload frmSelect
    
End Function


Function F_DadosPar(Optional ven As String = "", _
    Optional valor As String = "0,00") As parcela
    
    
    With frmDadosPar
        .txtbxValor_Che.Text = valor
        .txtbxVencimento.Text = ven
                
        On Error Resume Next
        
        .Show 1
        If .Tag <> "" Then
            
            F_DadosPar.valor = .txtbxValor_Che.Text
            F_DadosPar.vencimento = .txtbxVencimento.Text
            
        End If
    End With
        Unload frmDadosPar
        
    
    
End Function

