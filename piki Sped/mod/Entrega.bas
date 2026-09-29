Attribute VB_Name = "Entrega"
Function GravaEntrega(venda As String, entregador As String) As Boolean
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "insert into entrega(venda,saida,entregador,usuario,status)values("
    strSQL = strSQL & "'" & venda & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & entregador & "',"
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "1)"
    conn.Execute strSQL
    GravaEntrega = True
trata_Erro:
    Exit Function
End Function
Function EncaminhaEntregaEmLote(num_Ven() As String, entregador As String, hora_enc As String) As String
    Dim lote As String
    Dim i As Integer
    lote = MaxKeySQL("historico_entrega", "lote", "", "")
    For i = 1 To UBound(num_Ven())
        If Not EncaminhaEntrega(num_Ven(i - 1), entregador, hora_enc, lote) Then
            MsgBox "Erro na entrega em lote"
        End If
    Next
    EncaminhaEntregaEmLote = lote
    Exit Function
trata_Erro:

End Function

Function AlteraEntregador(id As String, entregador As String, DataHora As String) As Boolean

    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update entrega set "
    If DataHora <> "" Then
        strSQL = strSQL & "saida=" & F_GravaDataTempo(DataHora) & ","
    End If
    strSQL = strSQL & "entregador='" & entregador & "' "
    
    strSQL = strSQL & "where id=" & id
    conn.Execute strSQL
    AlteraEntregador = True
trata_Erro:
    Exit Function
    
End Function


Function GravaChegada(id As String, chegada As String) As Boolean

    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update entrega set "
    
    strSQL = strSQL & "chegada=" & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "status=2"
    
    strSQL = strSQL & "where id=" & id
    conn.Execute strSQL
    GravaChegada = True
trata_Erro:
    Exit Function
    
End Function

Function PreparaEntrega(num_Ven As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    PreparaEntrega = GravaHistoricoEntrega(num_Ven, strCod_Usu, "", "", "", Now, "", "", "")
    strSQL = "update " & TAB_VENDAS & " set entrega_ven=2 "
    strSQL = strSQL & " where num_Ven='" & num_Ven & "' "
    conn.Execute strSQL
    Exit Function
trata_Erro:
    
End Function
Function EncaminhaEntrega(num_Ven As String, entregador As String, hora As String, Optional lote As String = "") As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    EnviaEntrega = GravaHistoricoEntrega(num_Ven, "", strCod_Usu, entregador, "", "", hora, "", "", "", lote)
    strSQL = "update " & TAB_VENDAS & " set entrega_ven=3 "
    
    
    strSQL = strSQL & " where num_Ven='" & num_Ven & "' "
    conn.Execute strSQL
    EncaminhaEntrega = True
    Exit Function
trata_Erro:
    
End Function
Function concluientregaEmLote(lote As String, hora As String, obs As String) As Boolean
    Dim strSQL As String
    
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "update historico_entrega set "
    strSQL = strSQL & "confirmado_por='" & strCod_Usu & "',"
    strSQL = strSQL & "hora_conf =" & F_GravaDataTempo(hora)
    strSQL = strSQL & " where lote='" & lote & "' "
    conn.Execute strSQL
    
    strSQL = "select num_ven from historico_entrega where lote='" & lote & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update " & TAB_VENDAS & " set entrega_ven=4 "
            strSQL = strSQL & " where num_Ven='" & !num_Ven & "' "
            conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    
    concluientregaEmLote = True
    Exit Function
trata_Erro:
    
End Function
Sub RetiraVendaEntregaLote(num_Ven As String)
    Dim strSQL As String
    
    strSQL = "Update " & TAB_VENDAS & " set entrega_ven=2 where num_ven='" & num_Ven & "' "
    conn.Execute strSQL
    
    strSQL = "Update historico_entrega set lote=NULL where num_ven='" & num_Ven & "' "
    conn.Execute strSQL
End Sub

Function concluientrega(num_Ven As String, entregador As String, hora As String, obs As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    concluientrega = GravaHistoricoEntrega(num_Ven, "", "", "", strCod_Usu, "", "", "", hora, obs)
    strSQL = "update " & TAB_VENDAS & " set entrega_ven=4 "
    strSQL = strSQL & " where num_Ven='" & num_Ven & "' "
    conn.Execute strSQL
    
    
    
    concluientrega = True
    Exit Function
trata_Erro:
    
End Function

Function GravaHistoricoEntrega(num_Ven As String, Optional preparado_por As String = "", Optional encaminhado_pro As String = "", _
    Optional entregador As String = "", Optional confirmado_pro As String = "", Optional hora_prep As String = "", Optional hora_enc As String = "", _
    Optional hora_env As String = "", Optional hora_conf As String = "", Optional obs_entrega As String = "", Optional lote As String = "") As Boolean
    On Error GoTo trata_Erro
    Dim strSQL As String
    
    If F_BuscaUmCampo("num_ven", "historico_entrega", "num_ven='" & num_Ven & "'") = "" Then
        
        strSQL = "insert into historico_entrega(num_ven,preparado_por,encaminhado_por,entregador,confirmado_por,"
        strSQL = strSQL & "hora_prep, hora_enc,hora_env, hora_conf, obs_entrega)values("
        strSQL = strSQL & "'" & num_Ven & "',"
        strSQL = strSQL & "'" & preparado_por & "',"
        strSQL = strSQL & "'" & encaminhado_pro & "',"
        strSQL = strSQL & "'" & entregador & "',"
        strSQL = strSQL & "'" & confirmado_por & "',"
        If hora_prep = "" Then
            strSQL = strSQL & "NULL,"
        Else
            strSQL = strSQL & F_GravaDataTempo(hora_prep) & ","
        End If
        If hora_enc = "" Then
            strSQL = strSQL & "NULL,"
        Else
            strSQL = strSQL & F_GravaDataTempo(hora_enc) & ","
        End If
        If hora_env = "" Then
            strSQL = strSQL & "NULL,"
        Else
            strSQL = strSQL & F_GravaDataTempo(hora_env) & ","
        End If
        If hora_conf = "" Then
            strSQL = strSQL & "NULL,"
        Else
            strSQL = strSQL & F_GravaDataTempo(hora_conf) & ","
        End If
        If obs_entrega = "" Then
            strSQL = strSQL & "NULL)"
        Else
            strSQL = strSQL & "'" & obs_entrega & "',"
        End If
        
        conn.Execute strSQL
    Else
        strSQL = "update historico_entrega set "
        If preparado_por <> "" Then
            strSQL = strSQL & "preparado_por='" & preparado_por & "',"
        End If
        If encaminhado_pro <> "" Then
            strSQL = strSQL & "encaminhado_por='" & encaminhado_pro & "',"
        End If
        If entregador <> "" Then
            strSQL = strSQL & "entregador='" & entregador & "',"
        End If
        If confirmado_pro <> "" Then
            strSQL = strSQL & "confirmado_por='" & confirmado_pro & "',"
        End If
        If hora_prep <> "" Then
            strSQL = strSQL & "hora_prep=" & F_GravaDataTempo(hora_prep) & ","
        End If
        If hora_enc <> "" Then
            strSQL = strSQL & "hora_enc=" & F_GravaDataTempo(hora_enc) & ","
        End If
        If hora_env <> "" Then
            strSQL = strSQL & "hora_env=" & F_GravaDataTempo(hora_env) & ","
        End If
        If hora_conf <> "" Then
            strSQL = strSQL & "hora_conf=" & F_GravaDataTempo(hora_conf) & ", "
        End If
        If lote <> "" Then
            strSQL = strSQL & " lote='" & lote & "', "
        End If
        If obs_entrega <> "" Then
            strSQL = strSQL & " obs_entrega='" & obs_entrega & "', "
        End If
        strSQL = strSQL & " num_ven=num_ven "
        
       
        
        
        strSQL = strSQL & " where num_ven='" & num_Ven & "' "
        conn.Execute strSQL
    End If
    
    
    
    GravaHistoricoEntrega = True
trata_Erro:
    Exit Function
End Function

