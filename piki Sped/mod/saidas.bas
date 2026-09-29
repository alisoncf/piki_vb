Attribute VB_Name = "Saidas"
Type CFOP
    cfop_cfo As String
    nome_cfo As String
    tipo_cfo As String
End Type


Function F_GravaCupom(num_cup As String, num_ven As String, valor_cup As String, desconto_cup As String) As String
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    sSeq = ChaveSQL("cupons", "seq_cup", "")
    strSQL = "insert into cupons(seq_cup,num_cup,num_ven,valor_cup,desconto_cup,data_cup,"
    strSQL = strSQL & "cod_cai,cod_tur)values("
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & "'" & num_cup & "',"
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & F_GravaDec(valor_cup) & ","
    strSQL = strSQL & F_GravaDec(desconto_cup) & ","
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Cai & "',"
    strSQL = strSQL & "'" & strCod_Tur & "') "
    conn.Execute strSQL
    
    F_GravaCupom = sSeq
trata_Erro:
    
    
    
End Function
Function F_GravaItCup(seq_cup As String, seq_ite As String, cod_pro As String, _
    qtd_ite As String, unit_ite As String, Desconto_ite As String, icms_ite As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into tab_itcup (seq_cup, seq_ite, cod_pro, qtd_ite, unit_ite, desconto_ite,icms_ite ) VALUES ("
    strSQL = strSQL & seq_cup & ","
    strSQL = strSQL & seq_ite & ","
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(unit_ite) & ","
    strSQL = strSQL & F_GravaDec(Desconto_ite) & ","
    strSQL = strSQL & F_GravaDec(icms_ite) & ")"
    conn.Execute strSQL
    F_GravaItCup = True
trata_Erro:
    
End Function
Sub P_GravaCupom(num_ven As String)
    Dim strSQL As String
    Dim sCod As String
    Dim sUnit As String
    Dim sDesc As String
    Dim sQtd As String
    Dim iCount As Integer
    Dim rPesq As New ADODB.Recordset
    Dim cValor As Currency
    Dim sSeq As String
    Dim num_cup As String
    On Error GoTo trata_Erro
    Dim cICM As Currency
    Dim cTotalIcm As Currency
    strSQL = "select * from " & TAB_VENDAS & " where num_ven='" & num_ven & "' "
    rPesq.Open strSQL, conn_client
    
    iCount = ecf_NumeroCupom(num_cup)
    iCount = 0
    With rPesq
        sSeq = Saidas.F_GravaCupom(num_cup, num_ven, !valor_ven, !desconto_ven)
        .Close
    End With
    If sSeq = "" Then
        MsgBox "Erro"
    End If
    strSQL = "select distinct ti.*,perc_imp from " & TAB_ITVENDAS & " as ti," & TAB_PRODUTOS & " as tp, tab_impostos as tm "
    strSQL = strSQL & " where ti.cod_pro=tp.cod_pro and tm.cod_imp=tp.cod_imp and num_ven='" & num_ven & "' order by seq_ite "
    rPesq.Open strSQL, conn_client
    
    With rPesq
        Do While Not .EOF
            cICM = (!qtd_ite * !unit_ite * !perc_imp / 100)
            cTotalIcm = cTotalIcm + cICM
            If Saidas.F_GravaItCup(sSeq, .AbsolutePosition + 1, !cod_pro, !qtd_ite, !unit_ite, 0, CStr(cICM)) = True Then
                strSQL = "update " & TAB_PRODUTOS & " set ef_pro=ef_pro-" & F_GravaDec(!qtd_ite)
                strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
                conn.Execute strSQL
            Else
                MsgBox "Erro na gravação do ITCUP, favor avisar o responsável!", vbExclamation
            End If
            
            'baixar também a saída
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "update cupons set icms_cup=" & F_GravaDec(cTotalIcm)
    strSQL = strSQL & " where seq_cup='" & sSeq & "' "
    conn.Execute strSQL
Exit Sub
trata_Erro:
    On Error Resume Next
    Open "C:\erro.txt" For Output As #2
        Print #2, Err.Description & Chr(13) + Chr(10)
        Print #2, strSQL
    Close #2
    MsgBox "Ocorreu um erro na gravação, e foi gerado o arquivo c:\erro.txt", vbExclamation
    Exit Sub
End Sub


Sub P_GravaCFOP(cfop_cfo As String, nome_cfo As String, tipo_cfo As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tab_cfop(cfop_cfo,nome_cfo,tipo_cfo)values("
    strSQL = strSQL & "'" & cfop_cfo & "',"
    strSQL = strSQL & "'" & nome_cfo & "',"
    strSQL = strSQL & "'" & tipo_cfo & "')"
    conn.Execute strSQL
    Exit Sub
trata_Erro:
End Sub
Sub P_AlteraCFOP(cfop_cfo As String, nome_cfo As String, tipo_cfo As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update tab_cfop set  "
    
    strSQL = strSQL & "nome_cfo='" & nome_cfo & "',"
    strSQL = strSQL & "tipo_cfo='" & tipo_cfo & "' "
    strSQL = strSQL & " where cfop_cfo='" & cfop_cfo & "' "
    conn.Execute strSQL
    Exit Sub
trata_Erro:
End Sub

Sub P_DeleteCFOP(cfop_cfo As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "delete from tab_cfop where "
    strSQL = strSQL & "cfop_cfo='" & cfop_cfo & "' "
    
    conn.Execute strSQL
    Exit Sub
trata_Erro:
End Sub
Function F_SelectCFOP(Optional expr As String, Optional tipo As String = "S") As String
    Dim strSQL As String
    strSQL = "select * from tab_cfop where (nome_cfo like '" & expr & "%' "
    If expr <> "" Then
        strSQL = strSQL & " or cfop_cfo='" & expr & "'"
    End If
    
    strSQL = strSQL & ") "
    If tipo <> "" Then
        strSQL = strSQL & " and tipo_cfo='" & tipo & "' "
    End If
    F_SelectCFOP = strSQL
    
End Function
Function F_BuscaCFOP(expr As String, Optional tipo As String = "S") As CFOP
    frmEdiCfop.tag = "S"
    frmEdiCfop.caption = "Selecionar um CFOP"
    frmEdiCfop.txtbxExpr.Text = expr
    If tipo = "S" Then
        frmEdiCfop.optnSaida.Value = True
        frmEdiCfop.optnEntrada.Enabled = False
    End If
    frmEdiCfop.Show 1
    If frmEdiCfop.tag <> "" Then
        F_BuscaCFOP.cfop_cfo = frmEdiCfop.adc.Recordset!cfop_cfo
        F_BuscaCFOP.nome_cfo = frmEdiCfop.adc.Recordset!nome_cfo
        F_BuscaCFOP.tipo_cfo = frmEdiCfop.adc.Recordset!tipo_cfo
    End If
    Unload frmEdiCfop

End Function
Function F_BuscaCFOPCodigo(codigo As String) As CFOP
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_cfop where cfop_cfo='" & codigo & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount > 0 Then
            F_BuscaCFOPCodigo.cfop_cfo = !cfop_cfo
            F_BuscaCFOPCodigo.nome_cfo = !nome_cfo
            F_BuscaCFOPCodigo.tipo_cfo = !tipo_cfo
        End If
    End With
End Function

Sub P_ImprimeNota(obj As Object, Seq_sai As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rPro As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select tf.*,tc.*,ts.* from tab_saidasnf as ts, " & TAB_CLIENTES & " as tc, tab_cfop as tf "
    
    strSQL = strSQL & " where ts.cfop_sai=tf.cfop_cfo "
    strSQL = strSQL & " and tc.cod_cli=ts.cod_cli  "
    strSQL = strSQL & " and ts.seq_sai='" & Seq_sai & "'"
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Não encontrado!", vbExclamation
            .Close
            Exit Sub
        End If
        
        P_CabRelTimbr "NOTA FISCAL " & !nota_sai, obj
        obj.fontsize = 10
        P_ImprX obj, "Número: " & !nota_sai, 5
        P_ImprX obj, "Emissao: " & !emissao_sai, 35
        P_ImprX obj, "Operação: " & !cfop_cfo & " " & !nome_cfo, 70, False
        P_ImprX obj, "Cliente: " & !nome_Cli, 5
        P_ImprX obj, "CNPJ: " & !cpf_cli, 115
        P_ImprX obj, "IE: " & !rg_cli, 165, False
        
        P_line1 obj
        
        P_ImprX obj, "Valor: " & FormatVal(!valor_sai), 5
        P_ImprX obj, "Base de Cálculo: " & FormatVal(!bc_sai), 55
        P_ImprX obj, "ICMS: " & FormatVal(!vicms_sai), 115
        P_ImprX obj, "Produtos: " & FormatVal(!valp_sai), 155, False
        
        P_line1 obj
        
        strSQL = "select tp.nome_pro,tp.unidade_pro,tp.st_pro,ti.* from " & TAB_PRODUTOS & " as tp, tab_itsainf as ti  "
        strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
        strSQL = strSQL & " and ti.seq_sai='" & Seq_sai & "'"
        rPro.Open strSQL
        With rPro
            P_ImprX obj, "...", 5, True
            P_ImprX obj, "Código", 10, True
            P_ImprX obj, "Descrição", 30, True
            P_ImprX obj, "Unidade", 125, True
            P_ImprX obj, "Quant.", 150, True
            P_ImprX obj, "Unitário", 160, True
            P_ImprX obj, "Total", 180, False
            P_Line obj, 5, obj.CurrentY, 195, obj.CurrentY, vbDot
            
            Do While Not .EOF
                P_ImprX obj, .AbsolutePosition + 1, 5, True
                P_ImprX obj, !cod_pro, 10, True
                P_ImprX obj, !nome_pro, 30, True
                P_ImprX obj, !unidade_pro, 130, True
                P_ImprX obj, !qtd_its, 140, True
                P_ImprX obj, FormatVal(!unit_its), 160, True
                P_ImprX obj, FormatVal(!qtd_its * !unit_its), 180, False
                
                .MoveNext
            Loop
        End With
        
        
        
        
        
        .Close
    End With
    
    P_RodapeVis obj
    Exit Sub
trata_Erro:
    
MostraErro
End Sub
Sub imprimenota(Seq_sai As String)
    P_ImprimeNota frmVisualiza.Pic, Seq_sai
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_ImprimeNota Printer, Seq_sai
        Printer.EndDoc
        
    End If
    Unload frmVisualiza
End Sub
Sub P_InfoEntradasPro(cod_pro As String)
    Dim fr As New frmInfoEnt
    fr.tag = cod_pro
    fr.txtProduto.tag = cod_pro
    
    fr.Show 1
End Sub
Sub P_InfoSaidasPro(cod_pro As String)
    Dim fr As New frmInfSai
    fr.tag = cod_pro
    fr.Show 1

End Sub
