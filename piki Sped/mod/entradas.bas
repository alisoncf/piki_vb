Attribute VB_Name = "entradas"

Function EfetivaNotaF(seq_not As String, Optional avisa As Boolean = True) As Boolean
    Dim rPesq As New ADODB.Recordset
    Dim seq_ent As String
    
    Dim strSQL As String
    strSQL = "select * from tab_compras where seq_not='" & seq_not & "'"
    rPesq.Open strSQL, conn_client
    If avisa Then
        If rPesq.RecordCount = 0 Then
            MsgBox "Não encontrado!", vbExclamation
            rPesq.Close
            EfetivaNotaF = False
            Exit Function
            
        
        End If
        If rPesq!status_not >= 4 Then
            MsgBox "A nota fiscal já se encontra efetivada!", vbExclamation
            rPesq.Close
            EfetivaNotaF = False
            Exit Function
            
        
        End If
    End If
    rPesq.Close
    strSQL = "select ti.* from tab_compras as tc, tab_itcompras as ti "
    strSQL = strSQL & " where ti.seq_not=tc.seq_not "
    strSQL = strSQL & " and tc.seq_not='" & seq_not & "' "
    rPesq.Open strSQL, conn_client
    
    With rPesq
        
        Do While Not .EOF
            seq_ent = !seq_not
            strSQL = "update " & TAB_PRODUTOS & " set ef_pro=Ef_pro+" & F_GravaDec(!qtd_ite * CCur(!qtd_emb)) & ","
            strSQL = strSQL & "cf_pro=" & F_GravaDec(!custo_ite)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            conn.Execute strSQL
            
            
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "update tab_compras set status_not=4 where seq_not='" & seq_not & "'"
    conn.Execute strSQL
    
    
    EfetivaNotaF = True
    Exit Function
trata_erro:
    EfetivaNotaF = False
End Function
Function EfetivaNotaR(seq_not As String, Optional bEfetivaF As Boolean = True) As Boolean
    Dim rPesq As New ADODB.Recordset
    Dim seq_ent As String
    Dim cNovoPreco As Currency
    Dim strSQL As String
    Dim cEst As Currency
    Dim cCustoMed As Currency
    On Error GoTo trata_erro
    If strEfetnota_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Function
    End If
    
    strSQL = "select * from tab_cmp where seq_not='" & seq_not & "'"
    rPesq.Open strSQL, conn_client
    
    If rPesq.RecordCount = 0 Then
        MsgBox "Não encontrado!", vbExclamation
        Exit Function
    End If
    
    If Format(rPesq!seq_ent) <> "" Then
        If Not EfetivaNotaF(rPesq!seq_ent) Then
            MsgBox "Erro na efetivação da nota fiscal (F)", vbExclamation
            Exit Function
        End If
    End If
    rPesq.Close
    strSQL = "select tp.*,ti.* from tab_cmp as tc, tab_itcmp as ti," & TAB_PRODUTOS & " as tp"
    strSQL = strSQL & " where ti.seq_not=tc.seq_not "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tc.seq_not='" & seq_not & "'"
    rPesq.Open strSQL, conn_client
    
            
                
        

       
        With rPesq
            Do While Not .EOF
                cEst = !er_pro
                cCusto = !cr_pro
                
                If (cEst + !qtd_ite * CCur(!qtd_emb)) > 0 Then
                    cCustoMed = ((cEst * cCusto) + (!qtd_ite * CCur(!qtd_emb) * !custo_ite)) / (cEst + !qtd_ite * CCur(!qtd_emb))
                End If
                strSQL = "update " & TAB_PRODUTOS & " set er_pro=er_pro+" & F_GravaDec(!qtd_ite * CCur(!qtd_emb))
               
                
                strSQL = strSQL & ", cr_pro=" & F_GravaDec(!custo_ite)
                
                
                
                
                cNovoPreco = Round(!precofuturo_ite, 2)
                
                 
                strSQL = strSQL & ", pv_pro=" & F_GravaDec(cNovoPreco)
                    
                
                

                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                conn.Execute strSQL
                P_AltPreco !cod_pro, cNovoPreco, !pv_pro
                .MoveNext
            Loop
            
        End With
        
        strSQL = "update tab_cmp set status_not=4 where seq_not='" & seq_not & "'"
    conn.Execute strSQL
    
        
        
        EfetivaNotaR = True
    Exit Function
trata_erro:
    MostraErro
Exit Function
    
    
    
    
End Function
Sub ExcluirEntradasAntigas(data_maior As String, Optional operacao As String = "cmp")
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim t1 As String
    Dim t2 As String
    If operacao = "cmp" Then
        t1 = "tab_cmp"
        t2 = "tab_itcmp"
    Else
        t1 = "tab_compras"
        t2 = "tab_itcompras"
    End If
    
    strSQL = "select seq_not from " & t1
    strSQL = strSQL & " where  emissao_not <" & F_GravaData(data_maior)
    strSQL = strSQL & " order by seq_not "
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            strSQL = "delete from " & t2 & " where seq_not='" & !seq_not & "'"
            conn.Execute strSQL
            strSQL = "delete from " & t1 & " where seq_not='" & !seq_not & "'"
            conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
End Sub
