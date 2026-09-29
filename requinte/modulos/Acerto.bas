Attribute VB_Name = "Acerto"
Option Explicit
Function F_GravaAcertoAbertos(cod_Rep As String, data_ace As String, hora As String, _
    cod_cli As String, cod_For As String, valor_ace As String, cobranca_ace As String, _
    tipo_ter As String, comissao_ace As String, diferenca_ace As String, entrada_ace As String, status_ace As String, queima As String) As String
    Dim strSQL As String
    Dim sSeq As String
    If cod_For = "" Then
        cod_For = "01" 'dinheiro
    End If
    sSeq = MaxKeySQL("tab_acertosabertos", "seq_ace", "", "000000#")
    On Error GoTo trata_Erro
    strSQL = "insert into tab_acertosabertos(seq_ace,data_ace,hora_ace,cod_cli,cod_for,valor_ace,"
    strSQL = strSQL & "cod_usu,cobranca_ace,cod_rep,tipo_ter,comissao_ace,diferenca_ace,entrada_ace,ipusu_ace,queima)values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(data_ace) & ","
    strSQL = strSQL & F_GravaHora(hora, "HH:MM") & ","
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & "'" & cod_For & "',"
    strSQL = strSQL & F_GravaDec(valor_ace) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & cobranca_ace & "',"
    strSQL = strSQL & "'" & cod_Rep & "',"
    strSQL = strSQL & "'" & tipo_ter & "',"
    strSQL = strSQL & F_GravaDec(comissao_ace) & ","
    strSQL = strSQL & F_GravaDec(diferenca_ace) & ","
    strSQL = strSQL & F_GravaDec(entrada_ace) & ","
    strSQL = strSQL & "'" & strIP_Usu & "',"
    strSQL = strSQL & "'" & queima & "')"
    Conn.Execute strSQL
    F_GravaAcertoAbertos = sSeq
    
    Exit Function
trata_Erro:
    


End Function

Function F_BuscaDataUltAcerto(cod_Rep As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select max(data_ace) as data from tab_acertos where cod_rep='" & cod_Rep & "' and status_ace=2 "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        
        F_BuscaDataUltAcerto = !data
    End With
    Exit Function
trata_Erro:
    
    F_BuscaDataUltAcerto = "01/01/1900"
    

End Function



Sub P_AlteraAcertoAbertos(seq_ace As String, data_ace As String, hora As String, _
    cod_cli As String, cod_For As String, valor_ace As String, cobranca_ace As String, _
    tipo_ter As String, comissao_ace As String, diferenca_ace As String, entrada_ace As String, status_ace As String)
    Dim strSQL As String
    Dim sSeq As String
    If cod_For = "" Then
        cod_For = "01" 'dinheiro
    End If
    
    On Error GoTo trata_Erro
    strSQL = "update tab_acertosabertos set (seq_ace,data_ace,hora_ace,cod_cli,cod_for,valor_ace,"
    strSQL = strSQL & "cod_usu,cobranca_ace,tipo_ter,comissao_ace,diferenca_ace,entrada_ace,ipusu_ace)values("
    
    strSQL = strSQL & "data_ace=" & F_GravaData(data_ace) & ","
    strSQL = strSQL & "hora_ace=" & F_GravaHora(hora) & ","
    strSQL = strSQL & "cod_cli='" & cod_cli & "',"
    strSQL = strSQL & "cod_for='" & cod_For & "',"
    strSQL = strSQL & "valor_ace=" & F_GravaDec(valor_ace) & ","
    strSQL = strSQL & "cod_usu='" & strCod_Usu & "',"
    strSQL = strSQL & "cobranca_ace='" & cobranca_ace & "',"
    strSQL = strSQL & "tpo_ter='" & tipo_ter & "',"
    strSQL = strSQL & "comissao_ace=" & F_GravaDec(comissao_ace) & ","
    strSQL = strSQL & "diferenca_ace=" & F_GravaDec(diferenca_ace) & ","
    strSQL = strSQL & "entrada_ace=" & F_GravaDec(entrada_ace) & ","
    strSQL = strSQL & "ipusu_ace='" & strIP_Usu & "',"
    strSQL = strSQL & "status_ace='" & status_ace & "'"
    strSQL = strSQL & " where seq_ace='" & seq_ace & "' "
    Conn.Execute strSQL
    
    
    Exit Sub
trata_Erro:
    


End Sub


Function F_GravaAcerto(cod_Rep As String, data_ace As String, hora As String, _
    cod_cli As String, cod_For As String, valor_ace As String, cobranca_ace As String, _
    tipo_ter As String, comissao_ace As String, diferenca_ace As String, entrada_ace As String, status_ace As String, queima As String) As String
    Dim strSQL As String
    Dim sSeq As String
    If cod_For = "" Then
        cod_For = "01" 'dinheiro
    End If
    sSeq = MaxKeySQL("tab_acertos", "seq_ace", "", "000000#")
    On Error GoTo trata_Erro
    strSQL = "insert into tab_acertos(seq_ace,data_ace,hora_ace,cod_cli,cod_for,valor_ace,"
    strSQL = strSQL & "cod_usu,cobranca_ace,cod_rep,tipo_ter,comissao_ace,diferenca_ace,entrada_ace,status_ace,queima)values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(data_ace) & ","
    strSQL = strSQL & F_GravaHora(hora, "HH:MM") & ","
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & "'" & cod_For & "',"
    strSQL = strSQL & F_GravaDec(valor_ace) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & cobranca_ace & "',"
    strSQL = strSQL & "'" & cod_Rep & "',"
    strSQL = strSQL & "'" & tipo_ter & "',"
    strSQL = strSQL & F_GravaDec(comissao_ace) & ","
    strSQL = strSQL & F_GravaDec(diferenca_ace) & ","
    strSQL = strSQL & F_GravaDec(entrada_ace) & ","
    strSQL = strSQL & "'" & status_ace & "',"
    strSQL = strSQL & "'" & queima & "')"
    Conn.Execute strSQL
    F_GravaAcerto = sSeq
    
    Exit Function
trata_Erro:
    


End Function
Function F_GravaItAcerto(seq_ace As String, seq_con As String, cod_pro As String, unit_ite As String, qtd_ite As String, Optional dif_ite As String = "0") As String
    Dim strSQL As String
    Dim sSeq As String
    sSeq = MaxKeySQL("tab_itacertos", "seq_ite", "seq_ace='" & seq_ace & "'", "")
    On Error GoTo trata_Erro
    strSQL = "insert into tab_itacertos(seq_ace,seq_ite,seq_con,cod_pro,unit_ite,qtd_ite,dif_ite)values("
    strSQL = strSQL & "'" & seq_ace & "',"
    strSQL = strSQL & sSeq & ","
    
    strSQL = strSQL & "'" & seq_con & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(unit_ite) & ","
    strSQL = strSQL & F_GravaDec(qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(dif_ite) & ")"
    
    Conn.Execute strSQL
    F_GravaItAcerto = sSeq
    
    Exit Function
trata_Erro:
    


End Function
Function F_GravaItAcertoAberto(seq_ace As String, seq_con As String, cod_pro As String, unit_ite As String, qtd_ite As String, Optional dif_ite As String = "0") As String
    Dim strSQL As String
    Dim sSeq As String
    sSeq = MaxKeySQL("tab_itacertosabertos", "seq_ite", "seq_ace='" & seq_ace & "'", "")
    On Error GoTo trata_Erro
    strSQL = "insert into tab_itacertosabertos(seq_ace,seq_ite,seq_con,cod_pro,unit_ite,qtd_ite,dif_ite)values("
    strSQL = strSQL & "'" & seq_ace & "',"
    strSQL = strSQL & sSeq & ","
    
    strSQL = strSQL & "'" & seq_con & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(unit_ite) & ","
    strSQL = strSQL & F_GravaDec(qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(dif_ite) & ")"
    
    Conn.Execute strSQL
    
    'reservando o produto
    strSQL = "update tab_consignacoes set reservado_con=reservado_con+" & F_GravaDec(qtd_ite) & ""
    strSQL = strSQL & " where seq_con='" & seq_con & "' "
    strSQL = strSQL & " and cod_pro='" & cod_pro & "' "
    Conn.Execute strSQL
    
    
    
    F_GravaItAcertoAberto = sSeq
    
    Exit Function
trata_Erro:
    


End Function
Sub P_DeletaItAcertoAberto(seq_ace As String, seq_ite As String, qtd_ite As String, seq_con As String, cod_pro As String)


    Dim strSQL As String
    
    On Error GoTo trata_Erro
    strSQL = "delete from tab_itacertosabertos where seq_ace='" & seq_ace & "' "
    strSQL = strSQL & " and seq_ite=" & seq_ite
    Conn.Execute strSQL
    
    
    
    'retornando a reserva
    strSQL = "update tab_consignacoes set reservado_con=reservado_con-" & F_GravaDec(qtd_ite) & ""
    strSQL = strSQL & " where seq_con='" & seq_con & "' "
    strSQL = strSQL & " and cod_pro='" & cod_pro & "' "
    Conn.Execute strSQL
    
    
    
    
    
    Exit Sub
trata_Erro:
    


End Sub



Function F_GravaComissProv(cod_Rep As String, seq_ace As String, comissao As String, seq_Doc As String) As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    Dim strSQL As String
    sSeq = MaxKeySQL("tab_comissprov", "seq_com", "", "000000#")
    strSQL = "insert into tab_comissprov(seq_com,cod_rep,seq_ace,valor_com,reman_com,seq_doc)values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & "'" & cod_Rep & "',"
    strSQL = strSQL & "'" & seq_ace & "',"
    strSQL = strSQL & F_GravaDec(comissao) & ","
    strSQL = strSQL & F_GravaDec(comissao) & ","
    strSQL = strSQL & "'" & seq_Doc & "')"
    Conn.Execute strSQL
    F_GravaComissProv = sSeq
    Exit Function
trata_Erro:
    
    
End Function
Function F_GravaComissParc(seq_itd As String, cod_Rep As String, seq_ace As String, valor As String, _
    venc As String, tipo As String) As Boolean
    Dim sSeq As String
    On Error GoTo trata_Erro
    'tipo 1 = Itdoc 2 = Cheque
    Dim strSQL As String
    
    strSQL = "insert into tab_comissparc(seq_itd,cod_rep,seq_ace,valor_com,"
    strSQL = strSQL & "venc_com,status_com,tipo_com)values("
    strSQL = strSQL & "'" & seq_itd & "',"
    strSQL = strSQL & "'" & cod_Rep & "',"
    strSQL = strSQL & "'" & seq_ace & "',"
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & F_GravaData(venc) & ","
    strSQL = strSQL & "'1',"
    strSQL = strSQL & "'" & tipo & "')"
    Conn.Execute strSQL
    F_GravaComissParc = True
    Exit Function
trata_Erro:
    
P_MostraErro Err.Description
End Function
Function F_GravaComissParcCai(seq_itd As String, seq_cai As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into tab_comissparccai(seq_itd,seq_Cai)values("
    strSQL = strSQL & "'" & seq_itd & "',"
    strSQL = strSQL & "'" & seq_cai & "')"
    
    Conn.Execute strSQL
    F_GravaComissParcCai = True
    Exit Function
trata_Erro:
    
    
End Function



Sub P_PreencheTipoCob(cmbx As Object)
    P_Preenche cmbx, "tab_tipoCob", "concat(cod_tip,'-',nome_tip)", ""

End Sub

Sub P_VisualizaAcerto(seq_ace As String)
    frmVisualizarAcerto.tag = seq_ace
    frmVisualizarAcerto.Show 1
End Sub
Function F_SelectAcertosCli(cod_cli As String, Optional tipo_ter As String = "C", _
    Optional dataI As String = "", Optional dataF As String = "", _
        Optional sOrderBy As String = "seq_ace desc", Optional status_ace As String = "2") As String
    Dim strSQL As String
    strSQL = "select ta.*,tc.nome_cli,tr.nome_rep "
    strSQL = strSQL & " from tab_acertos as ta, tab_representantes as tr, "
    strSQL = strSQL & " tab_clientes as tc "
    strSQL = strSQL & " where ta.cod_cli=tc.cod_cli"
    strSQL = strSQL & " and ta.cod_rep=tr.cod_rep "
    strSQL = strSQL & " and ta.tipo_ter='" & tipo_ter & "' "
    If cod_cli <> "" Then
        strSQL = strSQL & " and ta.cod_cli='" & cod_cli & "' "
    End If
    If status_ace <> "" Then
        strSQL = strSQL & " and status_ace='" & status_ace & "' "
    End If
    F_SelectAcertosCli = strSQL
End Function
Function F_SelectItAcertosCli(cod_cli As String, Optional tipo_ter As String = "C", Optional dataI As String = "", Optional dataF As String = "", Optional sOrderBy As String = "seq_ace desc") As String
    Dim strSQL As String
    strSQL = "select ta.*,tp.nome_pro,ti.*,tc.nome_cli,tr.nome_rep "
    strSQL = strSQL & " from tab_acertos as ta, tab_representantes as tr, "
    strSQL = strSQL & " tab_clientes as tc,tab_produtos as tp, tab_itacertos as ti "
    strSQL = strSQL & " where ta.cod_cli=tc.cod_cli"
    strSQL = strSQL & " and ta.cod_rep=tr.cod_rep "
    strSQL = strSQL & " and ta.seq_ace=ti.seq_ace "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and status_ace=2 "
    strSQL = strSQL & " and ta.tipo_ter='" & tipo_ter & "' "
    If cod_cli <> "" Then
        strSQL = strSQL & " and ta.cod_cli='" & cod_cli & "' "
    End If
    
    F_SelectItAcertosCli = strSQL
End Function


Function F_DeletaAcerto(seq_ace As String) As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = F_BuscaUmCampo("count(*)", "tab_comissparc", "status_com=2 and seq_ace='" & seq_ace & "' ")
    If strSQL > "0" Then
        MsgBox "Este acerto não pode ser excluído pois existe comissão paga do mesmo!", vbExclamation
        F_DeletaAcerto = False
        Exit Function
    End If
    strSQL = F_BuscaUmCampo("count(*)", "tab_itdoc as ti, tab_doc as td", _
        "ti.seq_doc=td.seq_doc and reman_itd>valor_itd and num_doc='" & seq_ace & "' and esp_doc='A'")
    
    If strSQL > "0" Then
        MsgBox "Este acerto não pode ser excluído pois existem documentos já baixados!", vbExclamation
        F_DeletaAcerto = False
        Exit Function
    End If
    
    strSQL = "select ti.*,ta.cod_rep,ta.valor_ace,tf.tipo_cai  from tab_itacertos as ti,tab_acertos as ta, "
    strSQL = strSQL & " tab_formaspag as tf"
    strSQL = strSQL & " where ti.seq_ace=ta.seq_ace and ta.seq_ace='" & seq_ace & "' "
    strSQL = strSQL & " and ta.cod_for=tf.cod_for "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            strSQL = F_GravaCai("D", Date, CStr(!valor_ace), "CANCELAMENTO ACERTO " & !seq_ace, !seq_ace, !tipo_cai, _
            strCod_Cai, CStr(!valor_ace), CStr(!valor_ace), strCod_Cai, strCod_Tur, "000", "S")
        End If
        Do While Not .EOF
            If Not F_GravaEstCons(!cod_Rep, !cod_pro, "+" & !qtd_ite) Then
                MsgBox "Ocorreu um erro na volta do estoque, informe o responsável!", vbExclamation
            End If
                    
            If Not Produtos.F_AlteraEstRes(!cod_pro, "+" & !qtd_ite) Then
                MsgBox "Erro na baixa do estoque", vbExclamation
            End If
            strSQL = "update tab_consignacoes set reman_con=reman_con+" & F_GravaDec(!qtd_ite)
            strSQL = strSQL & " where seq_con='" & !seq_con & "'"
            strSQL = strSQL & " and cod_pro='" & !cod_pro & "'"
            Conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    strSQL = "select * from tab_doc where num_doc='" & seq_ace & "' "
    strSQL = strSQL & " and esp_doc='A' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "delete from tab_itdoc where seq_doc='" & !seq_Doc & "'"
            Conn.Execute strSQL
            strSQL = "delete from tab_cheques where seq_doc='" & !seq_Doc & "'"
            Conn.Execute strSQL
            .MoveNext
        Loop
        rPesq.Close
    End With
    
    strSQL = "delete from tab_itdoc where num_doc='" & seq_ace & "' "
    strSQL = strSQL & " and esp_doc='A' "
    
    strSQL = "delete from tab_acertofin where seq_ace='" & seq_ace & "' "
    'conn.execute strSQL
    
    strSQL = "delete from tab_venfin where num_ven='" & seq_ace & "' "
    strSQL = strSQL & " and operacao_fin='A' "
    Conn.Execute strSQL
    strSQL = "Update tab_acertos set status_ace=3 where seq_ace='" & seq_ace & "' "
    Conn.Execute strSQL
    
    
    strSQL = "delete from tab_comissparc where seq_ace='" & seq_ace & "' "
    Conn.Execute strSQL
    F_DeletaAcerto = True
    GravaAcertoCNC seq_ace
    
'
Exit Function
trata_Erro:
    P_MostraErro
End Function
Sub GravaAcertoCNC(seq_ace As String)
    Dim strSQL As String
    On Error GoTo trata_Erro:
    strSQL = "insert into tab_acertoscanc(seq_ace,data_can,cod_usu)values("
    strSQL = strSQL & "'" & seq_ace & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "')"
    Conn.Execute strSQL
    
    Exit Sub
trata_Erro:
    
End Sub
