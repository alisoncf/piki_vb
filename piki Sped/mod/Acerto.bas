Attribute VB_Name = "Acerto"
Option Explicit
Function F_GravaAcertoAbertos(cod_Rep As String, data_ace As String, hora As String, _
    cod_cli As String, cod_for As String, valor_ace As String, cobranca_ace As String, _
    tipo_ter As String, comissao_ace As String, diferenca_ace As String, entrada_ace As String, status_ace As String, queima As String) As String
    Dim strSQL As String
    Dim sSeq As String
    If cod_for = "" Then
        cod_for = "01" 'dinheiro
    End If
    sSeq = MaxKeySQL("tab_acertosabertos", "seq_ace", "", "000000#")
    On Error GoTo trata_ERro
    strSQL = "insert into tab_acertosabertos(seq_ace,data_ace,hora_ace,cod_cli,cod_for,valor_ace,"
    strSQL = strSQL & "cod_usu,cobranca_ace,cod_rep,tipo_ter,comissao_ace,diferenca_ace,entrada_ace,ipusu_ace,queima)values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(data_ace) & ","
    strSQL = strSQL & F_GravaHora(hora, "HH:MM") & ","
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & "'" & cod_for & "',"
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
    conn.Execute strSQL
    F_GravaAcertoAbertos = sSeq
    
    Exit Function
trata_ERro:
    


End Function

Sub P_AlteraAcertoAbertos(seq_ace As String, data_ace As String, hora As String, _
    cod_cli As String, cod_for As String, valor_ace As String, cobranca_ace As String, _
    tipo_ter As String, comissao_ace As String, diferenca_ace As String, entrada_ace As String, status_ace As String)
    Dim strSQL As String
    Dim sSeq As String
    If cod_for = "" Then
        cod_for = "01" 'dinheiro
    End If
    
    On Error GoTo trata_ERro
    strSQL = "update tab_acertosabertos set (seq_ace,data_ace,hora_ace,cod_cli,cod_for,valor_ace,"
    strSQL = strSQL & "cod_usu,cobranca_ace,tipo_ter,comissao_ace,diferenca_ace,entrada_ace,ipusu_ace)values("
    
    strSQL = strSQL & "data_ace=" & F_GravaData(data_ace) & ","
    strSQL = strSQL & "hora_ace=" & F_GravaHora(hora) & ","
    strSQL = strSQL & "cod_cli='" & cod_cli & "',"
    strSQL = strSQL & "cod_for='" & cod_for & "',"
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
    conn.Execute strSQL
    
    
    Exit Sub
trata_ERro:
    


End Sub


Function F_GravaAcerto(cod_Rep As String, data_ace As String, hora As String, _
    cod_cli As String, cod_for As String, valor_ace As String, cobranca_ace As String, _
    tipo_ter As String, comissao_ace As String, diferenca_ace As String, entrada_ace As String, status_ace As String, queima As String) As String
    Dim strSQL As String
    Dim sSeq As String
    If cod_for = "" Then
        cod_for = "01" 'dinheiro
    End If
    sSeq = MaxKeySQL("tab_acertos", "seq_ace", "", "000000#")
    On Error GoTo trata_ERro
    strSQL = "insert into tab_acertos(seq_ace,data_ace,hora_ace,cod_cli,cod_for,valor_ace,"
    strSQL = strSQL & "cod_usu,cobranca_ace,cod_rep,tipo_ter,comissao_ace,diferenca_ace,entrada_ace,status_ace,queima)values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(data_ace) & ","
    strSQL = strSQL & F_GravaHora(hora, "HH:MM") & ","
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & "'" & cod_for & "',"
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
    conn.Execute strSQL
    F_GravaAcerto = sSeq
    
    Exit Function
trata_ERro:
    


End Function
Function F_GravaItAcerto(seq_ace As String, seq_con As String, cod_pro As String, unit_ite As String, qtd_ite As String, Optional dif_ite As String = "0") As String
    Dim strSQL As String
    Dim sSeq As String
    sSeq = MaxKeySQL("tab_itacertos", "seq_ite", "seq_ace='" & seq_ace & "'", "")
    On Error GoTo trata_ERro
    strSQL = "insert into tab_itacertos(seq_ace,seq_ite,seq_con,cod_pro,unit_ite,qtd_ite,dif_ite)values("
    strSQL = strSQL & "'" & seq_ace & "',"
    strSQL = strSQL & sSeq & ","
    
    strSQL = strSQL & "'" & seq_con & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(unit_ite) & ","
    strSQL = strSQL & F_GravaDec(qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(dif_ite) & ")"
    
    conn.Execute strSQL
    F_GravaItAcerto = sSeq
    
    Exit Function
trata_ERro:
    


End Function
Function F_GravaItAcertoAberto(seq_ace As String, seq_con As String, cod_pro As String, unit_ite As String, qtd_ite As String, Optional dif_ite As String = "0") As String
    Dim strSQL As String
    Dim sSeq As String
    sSeq = MaxKeySQL("tab_itacertosabertos", "seq_ite", "seq_ace='" & seq_ace & "'", "")
    On Error GoTo trata_ERro
    strSQL = "insert into tab_itacertosabertos(seq_ace,seq_ite,seq_con,cod_pro,unit_ite,qtd_ite,dif_ite)values("
    strSQL = strSQL & "'" & seq_ace & "',"
    strSQL = strSQL & sSeq & ","
    
    strSQL = strSQL & "'" & seq_con & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(unit_ite) & ","
    strSQL = strSQL & F_GravaDec(qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(dif_ite) & ")"
    
    conn.Execute strSQL
    
    'reservando o produto
    strSQL = "update tab_consignacoes set reservado_con=reservado_con+" & F_GravaDec(qtd_ite) & ""
    strSQL = strSQL & " where seq_con='" & seq_con & "' "
    strSQL = strSQL & " and cod_pro='" & cod_pro & "' "
    conn.Execute strSQL
    
    
    
    F_GravaItAcertoAberto = sSeq
    
    Exit Function
trata_ERro:
    


End Function
Sub P_DeletaItAcertoAberto(seq_ace As String, seq_ite As String, qtd_ite As String, seq_con As String, cod_pro As String)


    Dim strSQL As String
    
    On Error GoTo trata_ERro
    strSQL = "delete from tab_itacertosabertos where seq_ace='" & seq_ace & "' "
    strSQL = strSQL & " and seq_ite=" & seq_ite
    conn.Execute strSQL
    
    
    
    'retornando a reserva
    strSQL = "update tab_consignacoes set reservado_con=reservado_con-" & F_GravaDec(qtd_ite) & ""
    strSQL = strSQL & " where seq_con='" & seq_con & "' "
    strSQL = strSQL & " and cod_pro='" & cod_pro & "' "
    conn.Execute strSQL
    
    
    
    
    
    Exit Sub
trata_ERro:
    


End Sub



Function F_GravaComissProv(cod_Rep As String, seq_ace As String, comissao As String, seq_doc As String) As String
    Dim sSeq As String
    On Error GoTo trata_ERro
    Dim strSQL As String
    sSeq = MaxKeySQL("tab_comissprov", "seq_com", "", "000000#")
    strSQL = "insert into tab_comissprov(seq_com,cod_rep,seq_ace,valor_com,reman_com,seq_doc)values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & "'" & cod_Rep & "',"
    strSQL = strSQL & "'" & seq_ace & "',"
    strSQL = strSQL & F_GravaDec(comissao) & ","
    strSQL = strSQL & F_GravaDec(comissao) & ","
    strSQL = strSQL & "'" & seq_doc & "')"
    conn.Execute strSQL
    F_GravaComissProv = sSeq
    Exit Function
trata_ERro:
    
    
End Function
Function F_GravaComissParc(seq_itd As String, cod_Rep As String, seq_ace As String, valor As String, _
    venc As String, tipo As String) As Boolean
    Dim sSeq As String
    On Error GoTo trata_ERro
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
    conn.Execute strSQL
    F_GravaComissParc = True
    Exit Function
trata_ERro:
    
P_MostraErro Err.Description
End Function
Function F_GravaComissParcCai(seq_itd As String, seq_cai As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_ERro
    
    strSQL = "insert into tab_comissparccai(seq_itd,seq_Cai)values("
    strSQL = strSQL & "'" & seq_itd & "',"
    strSQL = strSQL & "'" & seq_cai & "')"
    
    conn.Execute strSQL
    F_GravaComissParcCai = True
    Exit Function
trata_ERro:
    
    
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
    strSQL = strSQL & " " & TAB_CLIENTES & " as tc "
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
    strSQL = strSQL & " " & TAB_CLIENTES & " as tc," & TAB_PRODUTOS & " as tp, tab_itacertos as ti "
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
    On Error GoTo trata_ERro
    strSQL = F_BuscaUmCampo("count(*)", "tab_comissparc", "status_com=2 and seq_ace='" & seq_ace & "' ")
    If strSQL > "0" Then
        MsgBox "Este acerto não pode ser excluído pois existe comissão paga do mesmo!", vbExclamation
        F_DeletaAcerto = False
        Exit Function
    End If
    strSQL = F_BuscaUmCampo("count(*)", " " & TAB_ITDOC & "  as ti,  " & TAB_DOC & "  as td", _
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
            strCod_Cai, "0", "0", strCod_Cai, strCod_Tur, "000", "S")
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
            conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    strSQL = "select * from  " & TAB_DOC & "  where num_doc='" & seq_ace & "' "
    strSQL = strSQL & " and esp_doc='A' "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "delete from  " & TAB_ITDOC & "  where seq_doc='" & !seq_doc & "'"
            conn.Execute strSQL
            strSQL = "delete from tab_cheques where seq_doc='" & !seq_doc & "'"
            conn.Execute strSQL
            .MoveNext
        Loop
        rPesq.Close
    End With
    
    strSQL = "delete from  " & TAB_ITDOC & "  where num_doc='" & seq_ace & "' "
    strSQL = strSQL & " and esp_doc='A' "
    
    strSQL = "delete from tab_acertofin where seq_ace='" & seq_ace & "' "
    'conn.execute strSQL
    
    strSQL = "delete from tab_venfin where num_ven='" & seq_ace & "' "
    strSQL = strSQL & " and operacao_fin='A' "
    conn.Execute strSQL
    strSQL = "Update tab_acertos set status_ace=3 where seq_ace='" & seq_ace & "' "
    conn.Execute strSQL
    
    
    strSQL = "delete from tab_comissparc where seq_ace='" & seq_ace & "' "
    conn.Execute strSQL
    F_DeletaAcerto = True
    GravaAcertoCNC seq_ace
    
'
Exit Function
trata_ERro:
    P_MostraErro
End Function
Sub GravaAcertoCNC(seq_ace As String)
    Dim strSQL As String
    On Error GoTo trata_ERro:
    strSQL = "insert into tab_acertoscanc(seq_ace,data_can,cod_usu)values("
    strSQL = strSQL & "'" & seq_ace & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "')"
    conn.Execute strSQL
    
    Exit Sub
trata_ERro:
    
End Sub
Sub imprimeacerto(seq_ace As String)
    P_imprimeAcerto seq_ace, frmVisualiza.Pic
    frmVisualiza.Show 1
    
    If frmVisualiza.tag <> "" Then
            P_imprimeAcerto seq_ace, Printer
            Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub

Sub P_imprimeAcerto(seq_ace As String, obj As Object)

    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim iEsq As Single
    Dim sCob As String
    Dim sEntr As Currency
    Dim cTotal As Currency
    Dim sTipoTer As String
    
    sTipoTer = F_BuscaUmCampo("tipo_ter", "tab_acertos", "seq_ace='" & seq_ace & "'")
    If sTipoTer = "C" Then
    
    Else
    
    End If
    If sTipoTer = "C" Then
        strSQL = "select ta.status_ace,ta.entrada_ace,tc.nome_cli,tc.endereco_cli,tt.nome_tip,tc.bairro_cli,tc.cidade_cli,"
        strSQL = strSQL & "ta.comissao_ace,ta.diferenca_ace,ta.cobranca_ace,ta.data_ace,hora_ace,cod_for,valor_ace,cod_usu,cobranca_ace,"
        strSQL = strSQL & "ti.*,tp.nome_pro,tp.ref_pro from " & TAB_PRODUTOS & " as tp, tab_itacertos as ti,"
        strSQL = strSQL & "tab_acertos as ta, " & TAB_CLIENTES & " as tc, tab_tipocob as tt "
        strSQL = strSQL & " where tc.cod_cli=ta.cod_cli "
        strSQL = strSQL & " and ta.seq_ace=ti.seq_ace "
        strSQL = strSQL & " and tt.cod_tip=cobranca_ace "
        strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and ta.seq_ace='" & seq_ace & "' "
        strSQL = strSQL & " order by seq_ite "
    Else
        strSQL = "select ta.status_ace,tc.nome_rep as nome_cli,ta.entrada_ace,tt.nome_tip,endereco_rep as endereco_cli,bairro_rep as bairro_cli,cidade_rep as cidade_cli,"
        strSQL = strSQL & "ta.comissao_ace,ta.diferenca_ace,ta.cobranca_ace,ta.data_ace,hora_ace,cod_for,valor_ace,cod_usu,cobranca_ace,"
        strSQL = strSQL & "ti.*,tp.nome_pro,tp.ref_pro from " & TAB_PRODUTOS & " as tp, tab_itacertos as ti,"
        strSQL = strSQL & "tab_acertos as ta, tab_representantes as tc, tab_tipocob as tt "
        strSQL = strSQL & " where tc.cod_rep=ta.cod_cli "
        strSQL = strSQL & " and tt.cod_tip=cobranca_ace"
        strSQL = strSQL & " and ta.seq_ace=ti.seq_ace "
        strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and ta.seq_ace='" & seq_ace & "' "
        strSQL = strSQL & " order by seq_ite "

    End If
    rPesq.Open strSQL, conn
    iEsq = 1
    
    
    
    
    With rPesq
        If .RecordCount > 0 Then
            sEntr = !entrada_ace
            
            If !status_ace = "1" Then
                P_ImprX obj, "Este acerto encontra-se em aberto!", 0, False
            ElseIf !status_ace = "3" Then
                P_ImprX obj, "Este acerto foi cancelado!", 0, False
            End If
            
            obj.FontName = "times new roman"
            obj.fontsize = 12
            obj.Print
        
            sCob = !nome_tip
        
            P_ImprX obj, "Acerto:" & !seq_ace, 0 + iEsq, True
            obj.FontBold = True
            P_ImprX obj, StrConv(!nome_Cli, vbProperCase), 30 + iEsq, True
            obj.fontsize = 10
        
            P_ImprX obj, "Comissão " & FormatVal(!comissao_ace), 140 + iEsq, True, 3
            P_ImprX obj, "Dif: " & FormatVal(!diferenca_ace), 160 + iEsq, True, 3
        
            obj.fontsize = 16
        
            P_ImprX obj, "R$ " & FormatVal(!valor_ace), 195 + iEsq, False, 3
            obj.FontBold = False
            obj.FontName = "arial"
            P_line1 obj
            obj.fontsize = 10
            P_ImprX obj, "Produto:", 0 + iEsq
            P_ImprX obj, "Ref:", 100 + iEsq
            P_ImprX obj, "Qtd:", 120 + iEsq, True, True
            P_ImprX obj, "Unit:", 140 + iEsq, True, True
            P_ImprX obj, "Total:", 180 + iEsq, False, 3
            P_line1 obj, vbDot
        
            Do While Not .EOF
            
                If obj.CurrentY >= obj.ScaleHeight - 30 Then
                    P_RodapeVis obj
                    P_AddPicVis obj
                End If
            
            P_ImprX obj, StrConv(!nome_pro, vbProperCase), 0 + iEsq
            P_ImprX obj, !ref_pro, 100 + iEsq
            P_ImprX obj, !qtd_ite, 120 + iEsq, True, True
            P_ImprX obj, FormatVal(!unit_ite), 140 + iEsq, True, True
            P_ImprX obj, FormatVal(!unit_ite * !qtd_ite), 180 + iEsq, False, 3
            cTotal = cTotal + (!unit_ite * !qtd_ite)
            .MoveNext
        Loop
        P_line1 obj
        obj.FontBold = True
        P_ImprX obj, "Total..........:", 0 + iEsq
        P_ImprX obj, FormatVal(cTotal), 180 + iEsq, False, 3
        obj.FontBold = False
        P_line1 obj
        
    End If
        
        .Close
    End With
    
    
    
    
    strSQL = "select td.num_doc,ti.* from  " & TAB_ITDOC & "  as ti,  " & TAB_DOC & "  as td "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.esp_doc='A' "
    strSQL = strSQL & " and td.num_ven='" & seq_ace & "' "
    strSQL = strSQL & " and tipo_doc='R' "
    rPesq.Open strSQL, conn
    If rPesq.RecordCount > 0 Then
        obj.fontsize = 9
        obj.FontBold = True
        P_ImprX obj, "Duplicatas Geradas", 0 + iEsq, True
        P_ImprX obj, "Cobrança: " & sCob, 40 + iEsq, False, False
    
    End If
    obj.FontBold = False
    With rPesq
        strSQL = ""
        Do While Not .EOF
             If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj
                P_AddPicVis obj
             End If
             strSQL = strSQL & .AbsolutePosition & "ª Parc." & !venc_itd & " R$ " & FormatVal(!valor_itd) & "   "
            
            .MoveNext
            If Not .EOF Then strSQL = strSQL & ""
        Loop
        .Close
    End With
    P_ImprQuebra strSQL, obj, 0 + iEsq, obj.ScaleWidth - 10
    
    strSQL = "select td.num_doc,ti.num_che,valor_che,venc_che,cpf_che from tab_cheques as ti,  " & TAB_DOC & "  as td "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.esp_doc='A' "
    strSQL = strSQL & " and td.num_ven='" & seq_ace & "' "
    strSQL = strSQL & " and tipo_doc='R' "
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        obj.fontsize = 9
        obj.FontBold = True
        obj.Print
        P_ImprX obj, "Cheques Gerados", 0 + iEsq, False
        
    
    End If
    obj.FontBold = False
    With rPesq
        strSQL = ""
        Do While Not .EOF
             If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj
                P_AddPicVis obj
             End If
             strSQL = strSQL & "Ch: " & !num_che & " Venc: " & !Venc_che & "  R$ " & FormatVal(!Valor_che) & "   "
            
            .MoveNext
            If Not .EOF Then strSQL = strSQL & ""
        Loop
        .Close
    End With
    
    P_ImprQuebra strSQL, obj, 0 + iEsq, obj.ScaleWidth - 10
    
    P_ImprX obj, "", 0, False
    
    
    strSQL = "select tf.* from tab_caixa as tf,tab_acertofin as ta  "
    strSQL = strSQL & " where tf.seq_cai=ta.seq_fin "
    strSQL = strSQL & " and ta.seq_ace='" & seq_ace & "' "
    
    rPesq.Open strSQL, conn
    If rPesq.RecordCount > 0 Then
        obj.fontsize = 9
        obj.FontBold = True
        P_ImprX obj, "Pagamento Em Dinheiro", 0 + iEsq, False
        
    
    End If
    
    
    
    obj.FontBold = False
    With rPesq
        strSQL = ""
        Do While Not .EOF
             If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj
                P_AddPicVis obj
             End If
             strSQL = strSQL & "Valor " & FormatVal(!valor_Cai) & "   "
            
            .MoveNext
            If Not .EOF Then strSQL = strSQL & ""
        Loop
        .Close
    End With
    P_ImprQuebra strSQL, obj, 0 + iEsq, obj.ScaleWidth - 10
    
    
    
    strSQL = "select tv.*,tf.nome_fin from tab_venfin as tv, "
    strSQL = strSQL & "tab_financeiras as tf "
    strSQL = strSQL & " where tv.cod_fin=tf.cod_fin and tv.tipo_fin='A' "
    strSQL = strSQL & " and num_ven='" & seq_ace & "' "
    
    rPesq.Open strSQL, conn
    If rPesq.RecordCount > 0 Then
        obj.fontsize = 9
        obj.FontBold = True
        P_ImprX obj, "Pagamento Em Cartão", 0 + iEsq, False
        
    
    End If
    obj.FontBold = False
    With rPesq
        strSQL = ""
        Do While Not .EOF
             If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_RodapeVis obj
                P_AddPicVis obj
             End If
             strSQL = strSQL & "Cartão " & !nome_fin & " Valor " & FormatVal(!valor_fin) & "   "
            
            .MoveNext
            If Not .EOF Then strSQL = strSQL & ""
        Loop
        .Close
    End With
    P_ImprQuebra strSQL, obj, 0 + iEsq, obj.ScaleWidth - 10
    
    
    If sEntr > 0 Then
        P_ImprX obj, "Entrada: " & FormatVal(sEntr), 10, False
    End If
    
    
    
    
    
    obj.Print
    P_line1 obj
    
    'P_AddPicVis obj
End Sub

