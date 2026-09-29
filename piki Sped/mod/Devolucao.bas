Attribute VB_Name = "Devolucao"
Function GravaDevolucao(data_dev As String, hora_dev As String, cod_usu As String, valor_dev As String, _
    num_ven As String, Optional seq_dev As String = "", Optional motivo As String = "") As String

    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    If seq_dev = "" Then
        sSeq = MaxKeySQL("devolucoes", "seq_dev", "", "")
    Else
        sSeq = seq_dev
    End If
    strSQL = "insert into " & TAB_DEVOLUCOES & "(seq_dev,data_dev,hora_dev,cod_usu,valor_dev,num_ven,motivo)values("
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & F_GravaData(data_dev) & ","
    strSQL = strSQL & "'" & Format(hora_dev, "hh:mm") & "',"
    strSQL = strSQL & "'" & cod_usu & "',"
    strSQL = strSQL & F_GravaDec(valor_dev) & ","
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & "'" & motivo & "')"
    conn.Execute strSQL
    GravaDevolucao = sSeq
    Exit Function
    
trata_Erro:
    
End Function
Function AlteraMotivo(seq_dev As String, motivo As String) As Boolean

    Dim strSQL As String
    
    On Error GoTo trata_Erro
    
    strSQL = "update " & TAB_DEVOLUCOES
    strSQL = strSQL & " set motivo='" & motivo & "',"
    strSQL = strSQL & "cod_usu='" & strCod_Usu & "' "
    strSQL = strSQL & " where seq_dev='" & seq_dev & "' "
    conn.Execute strSQL
    AlteraMotivo = True
    Exit Function
    
trata_Erro:
    
End Function


Function GravaItDev(seq_dev As String, cod_pro As String, qtd_dev As String, total_dev As String, cod_emp As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into " & TAB_ITDEV & "(seq_dev,cod_pro,qtd_dev,total_dev,cod_emp)values("
    strSQL = strSQL & seq_dev & ","
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(qtd_dev) & ","
    strSQL = strSQL & F_GravaDec(total_dev) & ","
    strSQL = strSQL & "'" & cod_emp & "')"
    conn.Execute strSQL
    GravaItDev = True
    Exit Function
trata_Erro:
    
End Function

Sub ImprimeDevolucao(seq_dev As String)
    P_imprDevolucao seq_dev, frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_imprDevolucao seq_dev, Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub


Sub P_imprDevolucao(seq_dev As String, obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    Dim cliente As String
    On Error GoTo trata_Erro
    
    
    strSQL = "select distinct motivo, valor_dev,hora_dev,qtd_Dev,total_dev,abrevenda_usu,data_dev,nomecli_ven,dataefet_ven,(qtd_dev) as devolvido,tp.codfab_pro,tv.data_ven,dataefet_ven,"
    strSQL = strSQL & "tc.cnpj_cli,nome_cli,tc.cpf_cli,ti.* from " & TAB_VENDAS & " as tv "
    strSQL = strSQL & " inner join " & TAB_CLIENTES & " as tc "
    strSQL = strSQL & " on tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp on ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " inner join devolucoes as td on td.num_ven=tv.num_ven "
    strSQL = strSQL & " inner join itdev as tid on td.seq_dev=tid.seq_dev  and tid.cod_pro=ti.cod_pro "
    strSQL = strSQL & " where td.seq_dev='" & seq_dev & "' "
    
        
    
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Devolução não encontrada!", vbExclamation
            Exit Sub
        End If
        
        P_CabRelTimbr "DEVOLUÇÃO " & seq_dev, obj, True
        P_line1 obj
        obj.FontBold = True
        P_ImprX obj, "Data: " & !data_dev, 0
        P_ImprX obj, "Hora: " & !hora_dev, 30
        P_ImprX obj, "Venda: " & !num_ven, 50
        P_ImprX obj, "Data: " & !dataefet_ven, 80
        P_ImprX obj, "Vendedor: " & F_BuscaUmCampo("guerra_fun", TAB_FUNCIONARIOS, "cod_fun='" & !abrevenda_usu & "'"), 120, False
        P_ImprX obj, "Cliente: " & !nome_Cli, 0
        
        cliente = !nome_Cli
        P_ImprX obj, "Valor: " & FormatVal(!valor_dev), 120, False
        P_ImprX obj, "Motivo: " & TiraNull(!motivo), 0, False
        
        P_line1 obj
        obj.FontBold = False
        
        P_ImprX obj, "Ref.", 0
        P_ImprX obj, "Produto", 30
        P_ImprX obj, "EMP", 140
        P_ImprX obj, "Qtd Devolvida.", 150
        P_ImprX obj, "Total.", 180, False
        P_line1 obj, vbDot
        
        With rPesq
            Do While Not .EOF
                P_ImprX obj, !codfab_pro, 0
                P_ImprX obj, !desc_ite, 30
                P_ImprX obj, !cod_emp, 140
                P_ImprX obj, !qtd_dev, 150
                P_ImprX obj, FormatVal(!total_dev), 180, False
                
                .MoveNext
            Loop
            .Close
        End With
        P_line1 obj, vbDot
        
        obj.Print ""
        obj.Print ""
        
        P_CentraImpressao obj, String(Len(cliente) + 10, "_"), False
        P_CentraImpressao obj, cliente, False
        obj.Print ""
        P_line1 obj, vbDashDotDot
        
        
'        P_RodapeVis obj
        
        
    End With
    
    
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub


