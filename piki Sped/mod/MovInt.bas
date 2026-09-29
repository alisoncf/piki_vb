Attribute VB_Name = "MovInt"
Type DestinoMovInt
    cod_Des As String
    nome_des As String
End Type
Function F_GravaEspMovInt(nome_des As String) As String
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    If F_BuscaUmCampo("cod_des", "tab_destmovint", "nome_des='" & nome_des & "'") <> "" Then
        MsgBox "Já existe um destino com este nome!", vbExclamation
        Exit Function
    End If
    sCod = MaxKeySQL("tab_destmovint", "cod_des", "", "")
    
    
    
    strSQL = "Insert into tab_destmovint(cod_des,nome_des)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_des & "')"
    conn.Execute strSQL
    F_GravaEspMovInt = sCod
    
    Exit Function
trata_Erro:
End Function
Function F_AlteraEspMovInt(cod_Des As String, nome_des As String) As Boolean
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    If F_BuscaUmCampo("cod_des", "tab_destmovint", "nome_des='" & nome_des & "' and cod_Des<>" & cod_Des) <> "" Then
        MsgBox "Já existe um destino com este nome!", vbExclamation
        Exit Function
    End If
    
    
    
    
    strSQL = "update tab_destmovint set "
    
    strSQL = strSQL & "nome_des='" & nome_des & "' "
    strSQL = strSQL & " where cod_des=" & cod_Des
    conn.Execute strSQL
    F_AlteraEspMovInt = True
    
    Exit Function
trata_Erro:
End Function
Function F_SelectDest(Optional nome As String) As String
    Dim strSQL As String
    strSQL = "select nome_des as campo, cod_des as cod from tab_destmovint"
    strSQL = strSQL & " where (nome_des like '%" & nome & "%' or"
    strSQL = strSQL & " cod_des='" & nome & "') order by nome_des "
    F_SelectDest = strSQL
End Function
Function F_GeraNovoMovInt(sTipo As String, cod_Des As String, sDesc As String, valor As String, numero As String) As String
    Dim strSQL As String
    Dim sNum_ven As String
    sNum_ven = MaxKeySQL("tab_movint", "num_mov", "", "000000#")
    
    strSQL = "insert into tab_movint(Num_mov,tipo_mov,cod_des,desc_mov,valor_mov,"
    strSQL = strSQL & "data_mov,Situacao_mov,cod_usu,numero_mov)VALUES("
    
    strSQL = strSQL & "'" & sNum_ven & "', "
    strSQL = strSQL & "'" & sTipo & "', "
    strSQL = strSQL & "" & cod_Des & ", "
    strSQL = strSQL & "'" & sDesc & "', "
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & F_GravaData(Date) & ","
    
    strSQL = strSQL & "1,"
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & numero & "')"
    
    
    conn.Execute strSQL
    F_GeraNovoMovInt = sNum_ven
    
    
    
End Function
Function F_AlteraMovInt(num_mov As String, sTipo As String, cod_Des As String, sDesc As String, valor As String, situacao_mov As String) As Boolean
    Dim strSQL As String
    Dim sNum_ven As String
    
    On Error GoTo trata_Erro:
    strSQL = "update tab_movint set "
    strSQL = strSQL & "data_mov,Situacao_mov,cod_usu)VALUES("
    
    
    strSQL = strSQL & "tipo_mov='" & sTipo & "', "
    strSQL = strSQL & "cod_des=" & cod_Des & ", "
    strSQL = strSQL & "desc_mov='" & sDesc & "', "
    strSQL = strSQL & "valor_mov=" & F_GravaDec(valor) & ","
    strSQL = strSQL & "data_mov=" & F_GravaData(Date) & ","
    
    strSQL = strSQL & "situacao_mov='" & situacao_mov & "', "
    strSQL = strSQL & "cod_usu='" & strCod_Usu & "' "
    strSQL = strSQL & " where num_mov='" & num_mov & "' "
    
    
    conn.Execute strSQL
    F_AlteraMovInt = True
    
    
    Exit Function
trata_Erro:
    P_MostraErro
End Function

    
Function F_BuscaDestMovInt(desc As String) As DestinoMovInt
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select cod_des as cod,nome_des as nome from tab_destmovint "
    
    
    
    frmSelect.sSelect = strSQL
    
    Set frmSelect.cCriterio = Nothing
    frmSelect.cCriterio.Add "where cod_des>-1  "
    frmSelect.sOrd = " order by nome_des"
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.Msflxgrd.col = 0
        F_BuscaDestMovInt.cod_Des = frmSelect.Msflxgrd.Text
        frmSelect.Msflxgrd.col = 1
        F_BuscaDestMovInt.nome_des = frmSelect.Msflxgrd.Text
    End If
    Unload frmSelect

trata_Erro:
    P_MostraErro
End Function
Sub P_DeletaItMov(num_mov As String, seq_ite As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "delete from tab_itmov where num_mov='" & num_mov & "' "
    strSQL = strSQL & " and seq_ite=" & seq_ite
    conn.Execute strSQL
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Sub P_InsereItMov(num_mov As String, cod_pro As String, qtd_ite As String, custo_ite As String, unit_ite As String)
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    sSeq = MaxKeySQL("tab_itmov", "seq_ite", "num_mov='" & num_mov & "'", "")
    strSQL = "insert into tab_itmov(num_mov,cod_pro,qtd_ite,custo_ite,unit_ite,seq_ite)values("
    strSQL = strSQL & "'" & num_mov & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(custo_ite) & ","
    strSQL = strSQL & F_GravaDec(unit_ite) & ","
    strSQL = strSQL & sSeq & ") "
    conn.Execute strSQL
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub P_AlteraItMov(num_mov As String, seq_ite As String, cod_pro As String, qtd_ite As String, custo_ite As String, unit_ite As String)
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    
    strSQL = "update tab_itmov set "
    
    strSQL = strSQL & "cod_pro='" & cod_pro & "',"
    strSQL = strSQL & "qtd_ite=" & F_GravaDec(qtd_ite) & ","
    strSQL = strSQL & "custo_ite=" & F_GravaDec(custo_ite) & ","
    strSQL = strSQL & "unit_ite=" & F_GravaDec(unit_ite) & ""
    strSQL = strSQL & " where num_mov='" & num_mov & "' "
    strSQL = strSQL & " and seq_ite=" & seq_ite
    conn.Execute strSQL
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub P_GeraArquivoTransfMovInt(seq_mov As String, sCaminho As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If sCaminho = "" Then
        mdiFrm.cmndlg.FileName = "M" & seq_mov & ".txt"
        mdiFrm.cmndlg.Filter = ".txt"
        mdiFrm.cmndlg.ShowSave
        
        
        sCaminho = mdiFrm.cmndlg.FileName
        
    End If
    If sCaminho = "" Then
        sCaminho = App.Path & "M" & seq_mov
        MsgBox "O arquivo será salvo em " & sCaminho
    Else
        
    End If
    strSQL = F_SelectMovInt(seq_mov)
    rPesq.Open strSQL, conn
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Transferência " & seq_mov & " não encontrada!", vbExclamation
            .Close
            Exit Sub
        End If
        
        Open sCaminho For Output As #1
        Print #1, "0000"
        Print #1, "1000" & F_SoNumero(strCNPJ_Emp)
        Print #1, "1001" & !protocolo_mov
        Print #1, "1002" & !data_mov
        Print #1, "1003" & Now
        Print #1, "1004" & strCod_Usu
        Print #1, "1005" & strNome_Usu
        Do While Not .EOF
            Print #1, "2" & Format(.AbsolutePosition, "00#") & F_EsticaTexto(!cod_pro, 8);
            Print #1, F_EsticaTexto(!seq_ite, 3);
            Print #1, F_EsticaTexto(!qtd_ite, 8);
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
            Print #1, F_EsticaTexto(F_BuscaUmCampo("nome_for", TAB_FORNECEDORES, "cod_for='" & !cod_fab & "'"), 50)
            
            .MoveNext
        Loop
        Print #1, "9999"
        .Close
        Close #1
    End With
    MsgBox "Arquivo salvo " & sCaminho, vbInformation
    
End Sub
Function F_SelectMovInt(num_mov As String) As String
    Dim strSQL As String
    strSQL = "select tss.nome_sub,ts.cod_sec,ts.nome_sec,ti.*,td.nome_des,tm.*,ti.custo_ite*ti.qtd_ite as total,tp.* "
    strSQL = strSQL & " from tab_movint as tm, tab_itmov as ti,tab_destmovint as td,tab_secoes as ts, "
    strSQL = strSQL & TAB_SUBSEC & " as tss, "
    strSQL = strSQL & " " & TAB_PRODUTOS & " as tp where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and td.cod_des=tm.cod_des "
    strSQL = strSQL & " and tm.num_mov=ti.num_mov "
    strSQL = strSQL & " and tp.cod_sub=tss.cod_sub "
    strSQL = strSQL & " and ts.cod_sec=tss.cod_sec "
    strSQL = strSQL & " and tm.num_mov='" & num_mov & "' "
    F_SelectMovInt = strSQL
End Function

Function F_GeraNovoMovExt(cnpj_mov As String, sProtocolo As String) As String
    Dim strSQL As String
    Dim sNum_ven As String
    Dim sProtRec As String
    sProtRec = F_BuscaProtocolo
    sNum_ven = MaxKeySQL("tab_movext", "num_mov", "", "000000#")
    
    strSQL = "insert into tab_movext(Num_mov,cnpj_mov,protocolo_mov,"
    strSQL = strSQL & "data_mov,cod_usu,protocolorec_mov)VALUES("
    
    strSQL = strSQL & "'" & sNum_ven & "', "
    strSQL = strSQL & "'" & cnpj_mov & "', "
    strSQL = strSQL & "" & sProtocolo & ", "
    
    
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    
    
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & sProtRec & "')"
    
    
    conn.Execute strSQL
    F_GeraNovoMovExt = sNum_ven
    
    
    
End Function
Sub P_InsereItMovExt(num_mov As String, cod_pro As String, qtd_ite As String, cr_pro As String, _
    cf_pro As String, cb_pro As String, pv_pro As String)
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    sSeq = MaxKeySQL("tab_itmovext", "seq_ite", "num_mov='" & num_mov & "'", "")
    strSQL = "insert into tab_itmovext(num_mov,seq_ite,cod_pro,qtd_ite,cr_pro,cf_pro,cb_pro,pv_pro)values("
    strSQL = strSQL & "'" & num_mov & "',"
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(cr_pro) & ","
    strSQL = strSQL & F_GravaDec(cf_pro) & ","
    strSQL = strSQL & F_GravaDec(cb_pro) & ","
    strSQL = strSQL & F_GravaDec(pv_pro) & ")"
    conn.Execute strSQL
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Function F_BuscaProtocolo() As String
    Dim sProtocolo As String
    sProtocolo = ChaveSQL("tab_movint", "protocolo_mov", "#")
        
    sProtocolo = F_SoNumero(Format(Now, "yyyymmdd")) & sProtocolo
    F_BuscaProtocolo = sProtocolo
End Function
Sub P_CadDestMovInt()
    frmCadUmCampo.operacao = "5"
    frmCadUmCampo.Show 0
End Sub
Sub P_GeraArquivoProdutos(rPesq As Object, sCaminho As String)
    Dim strSQL As String
    Dim seq_mov As String
    
    
    seq_mov = F_SoNumero(Now) & F_SoNumero(strIP_Usu)
    
    
    If sCaminho = "" Then
        mdiFrm.cmndlg.FileName = "P" & seq_mov & ".txt"
        mdiFrm.cmndlg.Filter = ".txt"
        mdiFrm.cmndlg.ShowSave
        
        
        sCaminho = mdiFrm.cmndlg.FileName
        
    End If
    If sCaminho = "" Then
        sCaminho = App.Path & "M" & seq_mov
        MsgBox "O arquivo será salvo em " & sCaminho
    Else
        
    End If
    
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Não há produtos na lista!", vbExclamation
            .Close
            Exit Sub
        End If
        
        Open sCaminho For Output As #1
        Print #1, "0000"
        Print #1, "1000" & F_SoNumero(strCNPJ_Emp)
        Print #1, "1001" & seq_mov
        Print #1, "1002" & Date
        Print #1, "1003" & Now
        Print #1, "1004" & strCod_Usu
        Print #1, "1005" & strNome_Usu
        Do While Not .EOF
            Print #1, "2" & Format(.AbsolutePosition, "00#") & F_EsticaTexto(!cod_pro, 8);
            Print #1, F_EsticaTexto(!seq_ite, 3);
            Print #1, F_EsticaTexto(!qtd_ite, 8);
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
            Print #1, F_EsticaTexto(F_BuscaUmCampo("nome_for", TAB_FORNECEDORES, "cod_for='" & !cod_fab & "'"), 50)
            
            .MoveNext
        Loop
        Print #1, "9999"
        .Close
        Close #1
    End With
    MsgBox "Arquivo salvo " & sCaminho, vbInformation
    
End Sub

Sub ImprMovInt(num_mov As String)
    'imprimeMovInt frmVisualiza.Pic, num_mov
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
     '   imprimeMovInt Printer, num_mov
        Printer.EndDoc

    End If
            Unload frmVisualiza
End Sub

