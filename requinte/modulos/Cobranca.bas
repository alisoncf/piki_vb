Attribute VB_Name = "Cobranca"
'variáveis globais
Global strContaCob_Vendas As String
Global strContaCob_Avulso As String
Global strContaCob_AcertoLoja As String
Global strContaCob_AcertoRepr As String



Type Cobranca
    seq_cob As String
    data_cob As String
    entra_con As String
    
    resumo_cob As String
End Type
Type ContaCob
    cod_cob As String
    nome_cob As String
End Type
Type tipocob
    cod_tip As String
    nome_tip As String
    comiss_tip As String
End Type
Type tipopagfun
    cod_tip As String
    nome_tip As String
End Type

Type tipoCai
    cod_tip As String
    nome_tip As String
    
End Type


Function F_gravaCobranca(seq_ite As String, entra_con As String) As String
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    sSeq = MaxKeySQL("tab_cobrancas", "seq_cob", "", "000000#")
    strSQL = "insert into tab_cobrancas(seq_cob,data_cob,entra_con,hora_cob,seq_itd)values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & entra_con & "',"
    strSQL = strSQL & F_GravaHora(Time) & ","
    strSQL = strSQL & "'" & seq_ite & "') "
    Conn.Execute strSQL
    
    strSQL = "update tab_itdoc set cod_cob='" & entra_con & "' "
    strSQL = strSQL & " where seq_itd='" & seq_ite & "'"
    Conn.Execute strSQL
    
    F_gravaCobranca = sSeq
    
    
    Exit Function
    
trata_Erro:
    

End Function

Function F_AlteraCobranca(seq_cob As String, sai_con As String, datasai_cob As String, resumo_cob As String) As Boolean
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    
    strSQL = " update tab_cobrancas  set "
    strSQL = strSQL & " sai_con='" & sai_con & "',"
    
    strSQL = strSQL & "datasai_cob=" & F_GravaData(Now) & ","
    strSQL = strSQL & "resumo_cob='" & resumo_cob & "' "
    strSQL = strSQL & " where seq_cob='" & seq_cob & "' "
    Conn.Execute strSQL
    
    F_AlteraCobranca = True
    Exit Function
    
trata_Erro:
    
End Function
Function F_BuscaContaCob(expr As String) As ContaCob
    Dim strSQL As String
    If IsLoad("frmSelect") Then
        Unload frmSelect
    End If
    strSQL = "select cod_cob as cod,nome_cob as conta from tab_contacob where excluido_cob='N'"
    strSQL = strSQL & " and (cod_cob='" & expr & "' or nome_cob like '%" & expr & "%')"
    frmSelect.sOrd = "order by cod_cob "
    frmSelect.sSelect = strSQL
    frmSelect.cCriterio.Add " and 1=1 "
    
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        F_BuscaContaCob.cod_cob = frmSelect.msflxgrd.text
        frmSelect.msflxgrd.col = 1
        F_BuscaContaCob.nome_cob = frmSelect.msflxgrd.text
    End If
    Unload frmSelect
    
End Function
Function F_BuscaContaCobPorCod(cod_cob As String) As ContaCob
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_contacob "
    
    strSQL = strSQL & " where cod_cob='" & cod_cob & "' "
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            
            F_BuscaContaCobPorCod.cod_cob = !cod_cob
            
            F_BuscaContaCobPorCod.nome_cob = !nome_cob
        End If
        .Close
    End With
    
    
End Function

Function F_BuscaTipoCob(expr As String) As tipocob
    Dim strSQL As String
    strSQL = "select cod_tip as cod,nome_tip as tipo,comiss_tip as comiss from tab_Tipocob "
    frmSelect.sOrd = "order by nome_Tip"
    frmSelect.sSelect = strSQL
    frmSelect.cCriterio.Add " where cod_tip>0 "
    
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        F_BuscaTipoCob.cod_tip = frmSelect.msflxgrd.text
        frmSelect.msflxgrd.col = 1
        F_BuscaTipoCob.nome_tip = frmSelect.msflxgrd.text
        frmSelect.msflxgrd.col = 2
        F_BuscaTipoCob.comiss_tip = frmSelect.msflxgrd.text
    End If
    Unload frmSelect
    
End Function


Function F_GravaRepTip(cod_tip As String, cod_Rep As String, comissao_tip As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "delete from tab_reptip where cod_tip='" & cod_tip & "' "
    strSQL = strSQL & " and cod_rep='" & cod_Rep & "' "
    Conn.Execute strSQL
    
    strSQL = "insert into tab_reptip(cod_tip,cod_rep,comissao_tip)values("
    strSQL = strSQL & "'" & cod_tip & "',"
    strSQL = strSQL & "'" & cod_Rep & "',"
    strSQL = strSQL & F_GravaDec(comissao_tip) & ")"
    Conn.Execute strSQL
    F_GravaRepTip = True
    Exit Function
    P_MostraErro
trata_Erro:
    
End Function
Function DeletaRepTip(cod_tip As String, cod_Rep As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "delete from tab_reptip where cod_tip='" & cod_tip & "' "
    strSQL = strSQL & " and cod_rep='" & cod_Rep & "' "
    Conn.Execute strSQL
    
    
    DeletaRepTip = True
    Exit Function
    P_MostraErro
trata_Erro:
    
End Function
Function F_BuscaTipoCobRep(cod_Rep As String) As tipocob
    Dim strSQL As String
    strSQL = "select tt.cod_tip as cod,tt.nome_tip as tipo,tr.comissao_tip as comiss from tab_Tipocob as tt,"
    strSQL = strSQL & "tab_reptip as tr "
    frmSelect.sOrd = "order by nome_Tip"
    frmSelect.sSelect = strSQL
    frmSelect.cCriterio.Add " where tr.cod_tip=tt.cod_tip and cod_rep='" & cod_Rep & "' "
    
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        F_BuscaTipoCobRep.cod_tip = frmSelect.msflxgrd.text
        frmSelect.msflxgrd.col = 1
        F_BuscaTipoCobRep.nome_tip = frmSelect.msflxgrd.text
        frmSelect.msflxgrd.col = 2
        F_BuscaTipoCobRep.comiss_tip = frmSelect.msflxgrd.text
    End If
    
    Unload frmSelect
End Function
Function F_BuscaTipoPagFun() As tipopagfun
    Dim strSQL As String
    strSQL = "select cod_tip as cod,nome_tip as tipo from tab_tipopagfun "
    
    frmSelect.sOrd = "order by nome_Tip"
    frmSelect.sSelect = strSQL
    frmSelect.cCriterio.Add " where cod_tip>0 "
    
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        F_BuscaTipoPagFun.cod_tip = frmSelect.msflxgrd.text
        frmSelect.msflxgrd.col = 1
        F_BuscaTipoPagFun.nome_tip = frmSelect.msflxgrd.text
        
    End If
    Unload frmSelect
    
End Function
Function F_BuscaTipoPag() As tipopag
    Dim strSQL As String
    strSQL = "select cod_tip as cod,nome_tip as tipo from tab_tipopag "
    
    frmSelect.sOrd = "order by nome_Tip"
    frmSelect.sSelect = strSQL
    frmSelect.cCriterio.Add " where cod_tip>0 "
    
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        F_BuscaTipoPagFun.cod_tip = frmSelect.msflxgrd.text
        frmSelect.msflxgrd.col = 1
        F_BuscaTipoPagFun.nome_tip = frmSelect.msflxgrd.text
        
    End If
    Unload frmSelect
    
End Function

Function F_BuscaTipoCai() As tipoCai
    Dim strSQL As String
    strSQL = "select cod_tip as cod,nome_tip as tipo from tab_tipocai "
    
    frmSelect.sOrd = "order by nome_Tip"
    frmSelect.sSelect = strSQL
    frmSelect.cCriterio.Add " where cod_tip>0 "
    
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        F_BuscaTipoPagFun.cod_tip = frmSelect.msflxgrd.text
        frmSelect.msflxgrd.col = 1
        F_BuscaTipoPagFun.nome_tip = frmSelect.msflxgrd.text
        
    End If
    Unload frmSelect
    
End Function

