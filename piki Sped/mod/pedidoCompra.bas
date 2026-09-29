Attribute VB_Name = "pedidoCompra"
Function F_GravaPedidoCompra() As String
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    sCod = MaxKeySQL("tab_pedidocompra", "seq_ped", "", "")
    
    
    strSQL = "insert into tab_pedidocompra(seq_ped,data_ped,horaini_ped,cod_usu)values("
    strSQL = strSQL & sCod & ","
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & Format(Time, "hh:mm") & "',"
    strSQL = strSQL & "'" & strCod_Usu & "') "
    conn.Execute strSQL
    F_GravaPedidoCompra = sCod
    Exit Function
trata_Erro:
    
End Function

Sub P_AlteraPedidoCompra(seq_ped As String, horafin As String, total As String)
    Dim strSQL As String
        On Error GoTo trata_Erro
    
    
    strSQL = "update tab_pedidocompra  set "
    strSQL = strSQL & "horafin_ped='" & Format(Time, "hh:mm") & "',"
    strSQL = strSQL & "total_ped=" & F_GravaDec(total)
    strSQL = strSQL & " where seq_ped=" & seq_ped
    
    conn.Execute strSQL
    Exit Sub
trata_Erro:
    
End Sub
Sub P_GravaItPed(seq_ped As String, cod_pro As String, qtd_ite As String, unit_ite As String, cod_for As String, previsao_ite As String)
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    sSeq = MaxKeySQL("tab_itped", "seq_ite", "", "")
    
    strSQL = "insert into tab_itped(seq_ite,seq_ped,cod_for,previsao_ite,cod_pro,qtd_ite,unit_ite)values("
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & seq_ped & ","
    strSQL = strSQL & "'" & cod_for & "',"
    strSQL = strSQL & "'" & previsao_ite & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(unit_ite) & ")"
    
    conn.Execute strSQL
    
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Sub P_AlteraItPed(seq_ite As String, cod_pro As String, qtd_ite As String, unit_ite As String, cod_for As String)
    Dim strSQL As String
        On Error GoTo trata_Erro
        
    strSQL = "update tab_itped set "
    
    strSQL = strSQL & seq_ped & ","
    strSQL = strSQL & "cod_for='" & cod_for & "',"
    strSQL = strSQL & "cod_pro='" & cod_pro & "',"
    strSQL = strSQL & "qtd_ite=" & F_GravaDec(qtd_ite) & ","
    strSQL = strSQL & "unit_ite=" & F_GravaDec(unit_ite) & " "
    strSQL = strSQL & " where seq_ite=" & seq_ite
    conn.Execute strSQL
    
    
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Sub P_ExcluiItPed(seq_ite As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "delete from tab_itped where seq_ite=" & seq_ite
    conn.Execute strSQL
    Exit Sub
trata_Erro:
    MostraErro
    
End Sub
Function F_SelectItPed(seq_ped As String, Optional cod_for As String = "") As String
    Dim strSQL As String
    strSQL = "select tf.nome_for,tp.nome_pro,tp.unidade_pro,tp.er_pro,tp.cr_pro,ti.*,(qtd_ite*unit_ite) as total from "
    strSQL = strSQL & " " & TAB_PRODUTOS & " as tp, " & TAB_FORNECEDORES & " as tf,"
    strSQL = strSQL & " tab_itped as ti where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tf.cod_for=ti.cod_for "
    strSQL = strSQL & "  and seq_ped=" & seq_ped
    
    F_SelectItPed = strSQL
End Function
Function F_SelectPedidos(Optional seq_ped As String = "", Optional dataI As String = "", Optional dataF As String = "", _
    Optional cod_for As String = "", Optional cod_pro As String = "", Optional status As String = "1", _
    Optional orderBy As String = "data_ped desc") As String
    Dim strSQL As String
    
    strSQL = "select distinct tp.*,tf.guerra_fun from tab_pedidocompra as tp,tab_funcionarios as tf "
    If cod_for <> "" Or cod_pro <> "" Then
        strSQL = strSQL & " ,tab_itped as ti where ti.seq_ped=tp.seq_ped and "
    Else
        strSQL = strSQL & " where "
    End If
    strSQL = strSQL & " tf.cod_fun=tp.cod_usu "
    If dataI <> "" Then
        strSQL = strSQL & " and data_ped>=" & F_GravaDec(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and data_ped<=" & F_GravaDec(dataF)
    End If
    If cod_for <> "" Then
        strSQL = strSQL & " anc cod_for='" & cod_for & "'"
    End If
    If cod_prd <> "" Then
        strSQL = strSQL & " anc cod_pro='" & cod_pro & "'"
    End If
    If status <> "" Then
        strSQL = strSQL & " and status_ped='" & status & "'"
    End If
    If orderBy <> "" Then
        strSQL = strSQL & " order by " & orderBy
    End If
    F_SelectPedidos = strSQL
End Function

Sub ImprimePedidoCompra(seq_ped As String)
    Dim sArq As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rItens As New ADODB.Recordset
    strSQL = pedidoCompra.F_SelectPedidos(seq_ped)
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount = 0 Then
            Exit Sub
        End If
    End With
    sArq = html.F_AbreHtml("Pedido_" & rPesq!seq_ped)
    
    
    
    
    html.P_HtmlAbreHead 1
    html.P_HtmlAbreTitle 1, "Pedido " & seq_ped
    html.P_HtmlFechaHead 1
    html.P_HtmlAbreBody 1
    html.P_HtmlAbreH 1, 3, "PEDIDO DE COMPRA " & seq_ped, True
    html.P_HtmlAbreH 1, 4, "DATA: " & rPesq!data_ped, True
    strSQL = pedidoCompra.F_SelectItPed(seq_ped)
    rItens.Open strSQL
    P_rPesqToTable 1, rItens
    
    html.P_HtmlFechaBody 1
    html.P_FechaHtml 1, True, sArq
    
    
End Sub
