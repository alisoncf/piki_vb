Attribute VB_Name = "BancoDados"
Public conn As New ADODB.Connection
Public strConnect As String

Public Enum TipoRetornoIcms
    AliquotaNormal = 4
    AliquotaEcf = 3
    Descricao = 2
    IndiceEcf = 1
    cod_imp = 0
                
End Enum
Function AbreBancoAdo() As Boolean
    On Error GoTo trata_erro
    strConnect = "dsn=" & strNomeBanco
    conn.Open strConnect
    AbreBancoAdo = True
    Exit Function
trata_erro:
    P_MostraErro
End Function
Function FechaBancoAdo() As Boolean
    On Error Resume Next
    conn.Close
End Function
  Function F_GeraNovoOrc(sCod_Cli As String, nomecli As String, validade_orc As String, valor As String, desconto As String, condicoes As String) As String
    Dim strSQL As String
    Dim sNum_ven As String
    sNum_ven = MaxKeySQL("tab_Orc", "num_orc", "", "000000#")
    
    strSQL = "insert into tab_orc(Num_orc,Cod_Cli,Valor_orc,Desconto_orc,"
    strSQL = strSQL & "data_orc,Cod_Usu,Situacao_orc,nomecli_orc,validade_orc,obs_orc)values("
    
    strSQL = strSQL & "'" & sNum_ven & "', "
    strSQL = strSQL & "'" & sCod_Cli & "', "
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & F_GravaDec(desconto) & ","
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "1,"
    strSQL = strSQL & "'" & nomecli & "',"
    strSQL = strSQL & F_GravaData(validade_orc) & ","
    strSQL = strSQL & "'" & condicoes & "') "
    
    banco.Execute strSQL
    F_GeraNovoOrc = sNum_ven
    
    
    
End Function
Function F_GeraNovoMovInt(sTipo As String, sDesc As String, valor As String) As String
    Dim strSQL As String
    Dim sNum_ven As String
    sNum_ven = MaxKeySQL("tab_movint", "num_mov", "", "000000#")
    
    strSQL = "insert into tab_movint(Num_mov,tipo_mov,desc_mov,valor_mov,"
    strSQL = strSQL & "data_mov,Situacao_mov,cod_usu)VALUES("
    
    strSQL = strSQL & "'" & sNum_ven & "', "
    strSQL = strSQL & "'" & sTipo & "', "
    strSQL = strSQL & "'" & sDesc & "', "
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & F_GravaData(Date) & ","
    
    strSQL = strSQL & "1,"
    strSQL = strSQL & "'" & strCod_Usu & "')"
    
    
    banco.Execute strSQL
    F_GeraNovoMovInt = sNum_ven
    
    
    
End Function
 Function F_GeraNovaVenda(sCod_Cli As String, sCod_fun As String) As String
    Dim strSQL As String
    Dim sNum_ven As String
    sNum_ven = ChaveSQL("tab_vendas", "num_ven", "000000#")
    
    strSQL = "insert into tab_vendas(Num_Ven,Cod_Cli,Valor_Ven,Desconto_Ven,"
    strSQL = strSQL & "data_Ven,Cod_Fun,Cod_Usu,cod_tip,Situacao_Ven,nomecli_ven)values("
    
    strSQL = strSQL & "'" & sNum_ven & "', "
    strSQL = strSQL & "'" & sCod_Cli & "', "
    strSQL = strSQL & "0,0,"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "00,1,'VENDA A CONSUMIDOR') "
    
    banco.Execute strSQL
    F_GeraNovaVenda = sNum_ven
    
    
    
End Function
Function F_GeraNovaPerda(sCod_for As String, sCod_fun As String) As String
    Dim strSQL As String
    Dim sSeq_Per As String
    sSeq_Per = ChaveSQL("tab_perdas", "Seq_per", "000000#")
    strSQL = "insert into tab_perdas(cod_For,seq_per,data_per,cod_usu,situacao_per)"
    strSQL = strSQL & "values("
    strSQL = strSQL & "'" & sCod_for & "',"
    strSQL = strSQL & "'" & sSeq_Per & "',"
    strSQL = strSQL & F_GravaData(Date)
    strSQL = strSQL & ",'" & sCod_fun & "',"
    strSQL = strSQL & "'1')"
    banco.Execute strSQL
    

    F_GeraNovaPerda = sSeq_Per
    
    
    
End Function
Function F_GeraNovaTransf(sCod_Loj As String, sCod_fun As String) As String
    Dim strSQL As String
    Dim sSeq_Tra As String
    sSeq_Tra = ChaveSQL("tab_transf", "Seq_tra", "0000#")
    strSQL = "insert into tab_transf(cod_loj,seq_tra,data_Tra,cod_usu,situacao_tra)"
    strSQL = strSQL & "values("
    strSQL = strSQL & "'" & sCod_Loj & "',"
    strSQL = strSQL & "'" & sSeq_Tra & "',"
    strSQL = strSQL & F_GravaData(Date)
    strSQL = strSQL & ",'" & sCod_fun & "',"
    strSQL = strSQL & "'1')"
    banco.Execute strSQL
    

    F_GeraNovaTransf = sSeq_Tra
    
    
    
End Function

Function F_GeraNovoArranjo() As String
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_erro
    sSeq = ChaveSQL("tab_arranjos", "num_arr", "000000#")
    strSQL = "insert into tab_arranjos(num_arr,data_arr,hora_arr,cod_usu,situacao_arr)"
    strSQL = strSQL & "values("
    
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & F_GravaDataTempo(Date) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'1')"
    banco.Execute strSQL
    

    F_GeraNovoArranjo = sSeq
    Exit Function
trata_erro:
    F_GeraNovoArranjo = ""
    
    
End Function
Function F_GeraNovaDevolucao(sCod_Cli As String, sCod_fun As String) As String
    Dim strSQL As String
    Dim sNum_ven As String
    sNum_ven = ChaveSQL("tab_devolucoes", "num_dev", "000000#")
    strSQL = "insert into tab_devolucoes(Num_dev,Cod_Cli,Valor_dev,Desconto_dev,"
    strSQL = strSQL & "data_dev,Cod_Usu,cod_tip,Situacao_dev)values("
    
    strSQL = strSQL & "'" & sNum_ven & "', "
    strSQL = strSQL & "'" & sCod_Cli & "', "
    strSQL = strSQL & "0,0,"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "', "
    
    strSQL = strSQL & "00,1) "
    
    banco.Execute strSQL
    F_GeraNovaDevolucao = sNum_ven
    
    
    
End Function
Function F_GeraNovoPedido(sCod_Loj As String) As String
    Dim strSQL As String
    Dim sNum_ven As String
    sNum_ven = ChaveSQL("tab_pedidos", "num_ped", "000000#")
    strSQL = "insert into tab_pedidos(cod_loj,Num_Ped,data_ped,situacao_ped)values("
    strSQL = strSQL & "'" & sCod_Loj & "', "
    strSQL = strSQL & "'" & sNum_ven & "', "
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "1) "
    banco.Execute strSQL
    F_GeraNovoPedido = sNum_ven
    
    
    
End Function

Function F_GeraLog(sExp As String, sTabela As String) As String
    Dim strSQL As String
    
    strSQL = "insert into tab_log(cod_loj,expressao_log,data_log,status_log,seq_log,tab_log)Values("
    strSQL = strSQL & "'" & strCod_Loj & "',"
    strSQL = strSQL & "'" & F_AjustaTexto(sExp) & "',"
    strSQL = strSQL & "'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "',"
    strSQL = strSQL & "'0',"
    F_GeraLog = GeraMaxSQL("seq_log")
    strSQL = strSQL & F_GeraLog & ", "
    strSQL = strSQL & "'" & sTabela & "')"
    
    banco.Execute strSQL
End Function
Sub P_LogToText()
    Dim fs, a
    Dim strSQL As String
    Dim sNomeArq As String
    Dim rpesq As Recordset
    Dim strChunck As String
    Dim chTest
    P_Status "Gerando arquivo de Log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    sNomeArq = App.Path & "\" & strCod_Loj & Format(Now, "yyyymmddhhmmss") & ".stt"
    Set a = fs.CreateTextFile(sNomeArq, True)
    strSQL = "select * from tab_log where status_log=0"
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        If .RecordCount = 0 Then
            MsgBox "Não há Registros a serem transferidos no momento!", vbExclamation
            P_Status
            Exit Sub
        End If
        Do While Not .EOF
           
            'strChunck = .Fields(2).GetChunk(0, 32768)
             strChunck = F_BuscaUmCampo("expressao_log", "tab_log", "seq_log='" & !seq_log & "'")
            
            a.WriteLine !seq_log & Chr(9) & "'" & !cod_loj & "'" & Chr(9) & "'" & strChunck & "'" & _
                Chr(9) & "'" & Format(!data_log, "yyyy-mm-dd hh:mm:ss") & "'" & Chr(9) & _
                "'" & !status_log & "'" & Chr(9) & "'" & !tab_log & "'" & "|"
            strSQL = "update tab_log set status_log=1 where seq_log = '" & !seq_log & "' "
            banco.Execute strSQL
            .MoveNext
        Loop
        
        .Close
    End With
    
    a.Close
    P_Status
End Sub
Sub P_TextToTransf(sFile As String)
    Dim fs, a
    Dim s As String
    Dim strSQL As String
    Dim textLine As String
    Dim rpesq As Recordset
    Dim iCount As Integer
    Dim sAux As String
    Dim iPos As Long
    Dim iTam As Long
    Dim sERR As String
    On Error Resume Next
    P_Status "Gerando arquivo de Transferência..."
    Open sFile For Input As #1
    iTam = LOF(1)
    
    Do While Not EOF(1)
        
        Line Input #1, textLine
        For iCount = 1 To Len(textLine)
            
            iPos = iPos + 1
            s = Mid(textLine, iCount, 1)
            If s = "|" Then
                banco.Execute "insert into tab_transferencias values(" & sAux & ")"
                sAux = ""
                If Err.Description <> "" Then
                    sERR = sERR & Chr(13) + Chr(10)
                    Err.Clear
                End If
                
            ElseIf s = Chr(9) Then
                
                sAux = sAux & ","
                
            Else
                sAux = sAux & Mid(textLine, iCount, 1)
            End If
            
        Next
        P_Status "Byte " & iPos & " de " & iTam
    Loop

    If sERR <> "" Then
        MsgBox sERR, vbCritical, "Erros encontrados"
    End If
    Close #1
    
    P_Status
End Sub

Sub P_InsereEstoque(cod_pro As String, estoque_pro As String, custo_pro As String, _
    precomin_pro As String, preco_pro As String, cod_loj As String)
    Dim rpesq As Recordset
    Dim strSQL As String
    strSQL = "select cod_loj from tab_lojas order by cod_loj"
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        Do While Not .EOF
            If F_BuscaUmCampo("cod_pro", "tab_estoque", "cod_pro='" & cod_pro & _
            "' and cod_loj='" & !cod_loj & "'") = "" Then
                
                strSQL = "insert into tab_estoque(cod_pro,cod_loj,estoque_pro,"
                strSQL = strSQL & "custo_pro,customed_pro,ultcusto_pro,precomin_pro,"
                strSQL = strSQL & "preco_pro)Values("
                strSQL = strSQL & "'" & cod_pro & "', "
                strSQL = strSQL & "'" & !cod_loj & "', "
                If cod_loj = !cod_loj Then
                    strSQL = strSQL & F_GravaDec(estoque_pro) & ", "
                Else
                    strSQL = strSQL & "0, "
                End If
                strSQL = strSQL & F_GravaDec(custo_pro) & ", "
                strSQL = strSQL & F_GravaDec(custo_pro) & ", "
                strSQL = strSQL & F_GravaDec(custo_pro) & ", "
                strSQL = strSQL & F_GravaDec(precomin_pro) & ", "
                strSQL = strSQL & F_GravaDec(preco_pro) & ") "
                banco.Execute strSQL
            Else
                strSQL = "update tab_estoque set "
                If cod_loj = !cod_loj Then
                    strSQL = strSQL & "ESTOQUE_PRO=" & F_GravaDec(estoque_pro) & ", "
                End If
                strSQL = strSQL & "custo_PRO=" & F_GravaDec(custo_pro) & ", "
                strSQL = strSQL & "customed_PRO=" & F_GravaDec(custo_pro) & ", "
                strSQL = strSQL & "ultcusto_PRO=" & F_GravaDec(custo_pro) & ", "
                strSQL = strSQL & "precomin_PRO=" & F_GravaDec(precomin_pro) & ", "
                strSQL = strSQL & "preco_PRO=" & F_GravaDec(preco_pro) & " "
                strSQL = strSQL & " where cod_pro='" & cod_pro & "' "
                strSQL = strSQL & " and cod_loj='" & !cod_loj & "'"
                banco.Execute strSQL
            End If
            .MoveNext
        Loop
    End With
    
    
End Sub
Sub P_InsereEst_Loja(cod_pro As String, estoque_pro As String, custo_pro As String, _
    precomin_pro As String, preco_pro As String, cod_loj As String)
    Dim rpesq As Recordset
    Dim strSQL As String
    
    
    
    strSQL = "insert into tab_estoque(cod_pro,cod_loj,estoque_pro,"
    strSQL = strSQL & "custo_pro,customed_pro,ultcusto_pro,precomin_pro,"
    strSQL = strSQL & "preco_pro)Values("
    strSQL = strSQL & "'" & cod_pro & "', "
    strSQL = strSQL & "'" & cod_loj & "', "
    If cod_loj = F_BuscaUmCampo("cod_loj", "tab_parametros", "") Then
        strSQL = strSQL & F_GravaDec(F_BuscaUmCampo("Estoque_pro", "tab_produtos", _
                "cod_pro='" & cod_pro & "'")) & ", "
    Else
        strSQL = strSQL & "0, "
    End If
    strSQL = strSQL & F_GravaDec(custo_pro) & ", "
    strSQL = strSQL & F_GravaDec(custo_pro) & ", "
    strSQL = strSQL & F_GravaDec(custo_pro) & ", "
    strSQL = strSQL & F_GravaDec(precomin_pro) & ", "
    strSQL = strSQL & F_GravaDec(preco_pro) & ") "
    banco.Execute strSQL
        
    
    
    
End Sub

Function F_BuscaUmCampo(sCampoRetorno As String, sTabela As String, sWhere As String) As String
    Dim strSQL As String
    Dim rcrdstPesq As Recordset
    On Error GoTo trata_erro
    strSQL = "Select " & sCampoRetorno & " as Campo from " & sTabela
    If sWhere <> "" Then
        strSQL = strSQL & " Where " & sWhere
    End If
    Set rcrdstPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If rcrdstPesq.RecordCount < 1 Then
        F_BuscaUmCampo = ""
    Else
        F_BuscaUmCampo = rcrdstPesq!Campo
    End If
    rcrdstPesq.Close
Exit Function
trata_erro:
        F_BuscaUmCampo = ""
End Function
Function F_BuscaUmCampoOutroBanco(b As database, sCampoRetorno As String, sTabela As String, sWhere As String) As String
    Dim strSQL As String
    Dim rcrdstPesq As Recordset
    On Error GoTo trata_erro
    strSQL = "Select " & sCampoRetorno & " as Campo from " & sTabela
    If sWhere <> "" Then
        strSQL = strSQL & " Where " & sWhere
    End If
    Set rcrdstPesq = b.OpenRecordset(strSQL, dbOpenSnapshot)
    If rcrdstPesq.RecordCount < 1 Then
        F_BuscaUmCampoOutroBanco = ""
    Else
        F_BuscaUmCampoOutroBanco = rcrdstPesq!Campo
    End If
    rcrdstPesq.Close
Exit Function
trata_erro:
        F_BuscaUmCampoOutroBanco = ""
End Function

Function F_GravaDec(sValor) As String
    On Error GoTo trata_erro
    F_GravaDec = str(sValor)
    Exit Function
trata_erro:
    F_GravaDec = "0"
End Function
Function F_GravaHora(strHora) As String
Dim strDia, strMes, StrAno As String
  strHora = Format(strHora, "HH:MM:SS")
  
  F_GravaHora = "'" & strHora & "'"
  
  

End Function

Function F_GravaData(strData) As String
Dim strDia, strMes, StrAno As String
  strData = Format(strData, "DD/MM/YYYY")
  strDia = Mid(strData, 1, 2)
  strMes = Mid(strData, 4, 2)
  StrAno = Right(strData, 4)
  F_GravaData = "'" & StrAno & "-" & strMes & "-" & strDia & "'"
  If F_GravaData = "'--'" Then
    F_GravaData = "'0000-00-00'"
  End If
  

End Function
Function F_GravaDataTempo(strData) As String
Dim strDia, strMes, StrAno As String
  On Error GoTo trata_erro
  F_GravaDataTempo = "'" & Format(strData, "yyyy-mm-dd hh:mm:ss") & "' "
  Exit Function
trata_erro:
  F_GravaDataTempo = "'0000-00-00 00:00:00'"
End Function



Sub P_PreencheCateg(obj As Object)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select * from tab_categorias order by nome_Cat "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            obj.AddItem rpesq!nome_cat
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub
Sub P_PreencheMarca(obj As Object)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select * from tab_marcas order by nome_mar "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            obj.AddItem rpesq!nome_mar
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_PreencheFinanceiras(obj As Object, Optional tipo As String = "")
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select * from tab_financeiras "
    If tipo <> "" Then
        strSQL = strSQL & " where tipo_fin='" & tipo & "' "
    End If
    strSQL = strSQL & " order by nome_fin "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            obj.AddItem rpesq!nome_fin
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub


Sub P_PreencheFormasPag(obj As Object, Optional tipo_tip As String = "C")
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset
    'd - para duplicata, c - para cheque
    strSQL = "select * from tab_TIPOVEN  where cod_tip <> '00'"
    If tipo_tip <> "" Then
        strSQL = strSQL & " and tipo_tip='" & tipo_tip & "' "
    End If
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            obj.AddItem rpesq!cod_tip & "-" & rpesq!Descricao_tip
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub
Sub P_PreencheVendedor(obj As Object, Optional opcao As Integer = 0)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset 'se opcao = 0 traz nome completo , se 1 nome de guerra
    strSQL = "select tf.* from tab_funcionarios as tf, tab_cargos as tc "
    strSQL = strSQL & " where tc.cod_car=tf.cod_car and tc.venda_car='S' "
    strSQL = strSQL & " and tf.status_fun='1' "
    If opcao = 0 Then
        strSQL = strSQL & " order by nome_fun "
    Else
        strSQL = strSQL & " order by guerra_fun"
    End If
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            If opcao = 0 Then
                obj.AddItem rpesq!nome_fun
            Else
                obj.AddItem rpesq!guerra_fun
            End If
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_PreencheBanco(obj As Object)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select * from tab_Bancos order by Cod_Bco "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            obj.AddItem rpesq!cod_bco & "-" & rpesq!nome_bco
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub
Sub P_PreencheTurnos(obj As Object)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select * from tab_turnos order by Cod_tur "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            obj.AddItem rpesq!cod_tur & " - " & rpesq!horaini_tur & " às " & rpesq!horafin_tur
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_PreencheCaixas(obj As Object)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select * from tab_caixas order by Cod_cai "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            obj.AddItem rpesq!Cod_Cai & "-" & rpesq!Nome_cai
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub
Sub P_PreencheFornecedores(obj As Object)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select * from tab_fornecedores order by nome_for "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            obj.AddItem rpesq!nome_for
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_PreencheCargos(obj As Object)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select * from tab_cargos order by nome_car "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            obj.AddItem rpesq!nome_car
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_PreencheContas(obj As Object, Optional bFlag As Byte = 1, Optional bCaixa As Boolean = True)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset
    'Flag: 0 - só o código, 1 - código + nome, 2 - só nome
    'bcaixa:  se true carrega conta 001(caixa),se não oculta
    strSQL = "select * from tab_Contas "
    If bCaixa = False Then
        strSQL = strSQL & " Where cod_con <>  '001'"
    End If
    If bFlag = 0 Or bFlag = 2 Then
        strSQL = strSQL & "order by cod_con "
    
    Else
        strSQL = strSQL & "order by nome_con "
    
    End If
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            If bFlag = 0 Or bFlag = 2 Then
                obj.AddItem rpesq!cod_con
            ElseIf bFlag = 1 Then
                obj.AddItem rpesq!cod_con & "-" & rpesq!nome_con
            Else
                obj.AddItem rpesq!nome_con
            End If
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub


Sub P_PreencheImpostos(obj As Object)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select * from tab_impostos "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            obj.AddItem rpesq!cod_imp & "-" & rpesq!desc_imp
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub



Sub P_PreencheFunc(obj As Object, Optional opcao As Integer = 0)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset 'se opcao = 0 traz nome completo , se 1 nome de guerra
    strSQL = "select * from tab_funcionarios "
    If opcao = 0 Then
        strSQL = strSQL & " order by nome_fun "
    Else
        strSQL = strSQL & " order by guerra_fun"
    End If
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            If opcao = 0 Then
                obj.AddItem rpesq!nome_fun
            Else
                obj.AddItem rpesq!guerra_fun
            End If
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub


Function F_GravaItDoc(seq_doc As String, Valor_Itd As String, Venc_Itd As String, Parc_Itd) As String
            
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "insert into tab_itdoc1(Seq_Doc,Valor_Itd,Venc_Itd,Reman_Itd,Seq_Itd,Parc_Itd)Values( "
    
    strSQL = strSQL & "'" & seq_doc & "',"
    strSQL = strSQL & F_GravaDec(Valor_Itd) & ","
    strSQL = strSQL & F_GravaData(Venc_Itd) & ","
    strSQL = strSQL & F_GravaDec(Valor_Itd) & ","
    
    F_GravaItDoc = MaxKeySQL("tab_itdoc1", "seq_itd", "", "000000#")
    strSQL = strSQL & "'" & F_GravaItDoc & "', "
    strSQL = strSQL & Parc_Itd & ")"
    banco.Execute strSQL
    Exit Function
trata_erro:
    MsgBox Err.Description, vbCritical
    
End Function


Function F_GravaVenFin(num_ven As String, Cod_Fin As String, valor_ven As Currency, _
    data_ven As String, cod_usu As String, cod_cli As String, autorizacao_ven As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_erro
    F_GravaVenFin = True
    strSQL = "insert into tab_venfin(num_ven,cod_fin,valor_ven,data_ven,cod_usu,cod_cli,"
    strSQL = strSQL & "autorizacao_ven)values("
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & "'" & Cod_Fin & "',"
    strSQL = strSQL & F_GravaDec(valor_ven) & ","
    strSQL = strSQL & F_GravaData(data_ven) & ","
    strSQL = strSQL & "'" & cod_usu & "' ,"
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & "'" & autorizacao_ven & "') "
    banco.Execute strSQL
    Exit Function
trata_erro:
    MsgBox Err.Description, vbCritical
    F_GravaVenFin = False
End Function
Function F_GravaDoc(hist_doc As String, num_doc As String, Valor_Doc As String, Data_Doc As String, _
            cod_ter As String, tipo_doc As String, num_ven As String, _
            Parc_Doc As String, cod_pla As String) As String
    On Error GoTo trata_erro
    Dim strSQL As String
    
    If cod_pla = "" Then
        If tipo_doc = "R" Then
            cod_pla = strPlaNaoClassC
        Else
            cod_pla = strPlaNaoClassD
        End If
        
    End If
    
    strSQL = "insert into tab_doc1(historico_doc, Num_Doc,Valor_doc,Data_Doc,cod_ter, "
    strSQL = strSQL & "tipo_doc,num_ven,parc_doc,seq_doc,cod_pla)Values("
    strSQL = strSQL & "'" & hist_doc & "',"
    strSQL = strSQL & "'" & num_doc & "',"
    strSQL = strSQL & F_GravaDec(Valor_Doc) & ","
    strSQL = strSQL & F_GravaData(Data_Doc) & ","
    strSQL = strSQL & "'" & cod_ter & "',"
    strSQL = strSQL & "'" & tipo_doc & "',"
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & Parc_Doc & ","
    F_GravaDoc = ChaveSQL("tab_doc1", "seq_doc", "000000#")
    strSQL = strSQL & "'" & F_GravaDoc & "', "
    strSQL = strSQL & "'" & cod_pla & "') "
    banco.Execute strSQL
    Exit Function
trata_erro:
    F_GravaDoc = ""
    
End Function
Function MaxKeySQL(sTab As String, sCampo As String, sWh As String, formato As String) As String
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select max(" & sCampo & ") as campo  from " & sTab
    If sWh <> "" Then
        strSQL = strSQL & " where " & sWh
    End If
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        If IsNull(!Campo) Then
            MaxKeySQL = Format(1, formato)
        Else
            MaxKeySQL = Format(F_SoNumero(!Campo) + 1, formato)
        End If
    End With
    rpesq.Close
End Function
Function GeraMaxSQL(sCampo As String, Optional formato As String = "") As String
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select " & sCampo & " as campo  from tab_keys "
    
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        GeraMaxSQL = Format(!Campo + 1, formato)
        strSQL = "update tab_keys set " & sCampo & "=" & GeraMaxSQL
        banco.Execute strSQL
        
        
    End With
    rpesq.Close
End Function
Function ChaveSQL(sTable As String, sField As String, formato As String) As String
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select value_cod  as campo  from tab_codigos "
    strSQL = strSQL & " where table_cod='" & sTable & "' "
    strSQL = strSQL & " and field_cod='" & sField & "' "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        If IsNull(!Campo) Or .RecordCount = 0 Then
            strSQL = "insert into tab_codigos(table_cod,field_cod,value_cod)values("
            strSQL = strSQL & "'" & sTable & "',"
            strSQL = strSQL & "'" & sField & "',"
            strSQL = strSQL & "'" & Format("1", formato) & "')  "
            banco.Execute strSQL
            ChaveSQL = Format("1", formato)
        Else
            
            ChaveSQL = Format(!Campo + 1, formato)
            strSQL = "update tab_codigos set value_cod='" & ChaveSQL & "' "
            strSQL = strSQL & " where table_cod='" & sTable & "' "
            strSQL = strSQL & " and field_cod='" & sField & "' "
            banco.Execute strSQL
        End If
    End With
    rpesq.Close
End Function
Function F_GravaCHV(seq_cai As String, num_che As String, cod_bco As String, _
    age_che As String, cco_che As String, cpf_che As String, nome_che As String, Valor_che As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_erro
    F_GravaCHV = True
    strSQL = "insert into tab_chv(seq_cai,num_che,cod_bco,age_che,cco_che,cpf_che,"
    strSQL = strSQL & "nome_Che,valor_che)values("
    strSQL = strSQL & "'" & seq_cai & "',"
    strSQL = strSQL & "'" & num_che & "',"
    strSQL = strSQL & "'" & cod_bco & "',"
    strSQL = strSQL & "'" & age_che & "',"
    strSQL = strSQL & "'" & cco_che & "',"
    strSQL = strSQL & "'" & cpf_che & "',"
    strSQL = strSQL & "'" & nome_che & "',"
    strSQL = strSQL & F_GravaDec(Valor_che) & ")"
    banco.Execute strSQL
    Exit Function
trata_erro:
    F_GravaCHV = False
End Function


Function F_GravaCai(tipo_cai As String, data_cai As String, _
    valor_Cai As String, historico_cai As String, num_ven As String, _
    Esp_cai As String, cod_con As String, entrada_cai As Currency, troco_cai As Currency, caixa As String, turno As String) As String
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim sSeq  As String
    If tipo_cai = "R" Or tipo_cai = "+" Then
        tipo_cai = "C"
    End If
    If tipo_cai = "P" Or tipo_cai = "-" Then
        tipo_cai = "D"
    End If
    If Esp_cai = "1" Then
        Esp_cai = "DIN"
    End If
    'If num_ven <> "" Then
        'sSeq = F_BuscaUmCampo("seq_cai", "tab_caixa", "num_ven='" & num_ven & "' " & _
            " and cod_cai='" & CAIXA & "' and cod_Tur=" & turno & _
            " and Esp_cai='" & Esp_cai & "' " & _
            " and valor_cai=" & F_GravaDec(valor_Cai))
        'If sSeq <> "" Then
        '    MsgBox "Esta venda está duplicada, haverá a vinculação com o lançamento anterior!", vbInformation
        '    F_GravaCai = sSeq
        '    Exit Function
        'End If
    'End If
    strSQL = "insert into tab_caixa(seq_cai,data_cai,valor_Cai,historico_Cai,tipo_Cai,"
    strSQL = strSQL & "num_ven,esp_cai,cod_con,entrada_Cai,troco_cai,ip_cai,cod_cai,cod_tur,cod_usu,hora_cai)Values("
    sSeq = ChaveSQL("tab_caixa", "seq_cai", "000000#")
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(data_cai) & ","
    strSQL = strSQL & F_GravaDec(valor_Cai) & ","
    strSQL = strSQL & "'" & historico_cai & "',"
    strSQL = strSQL & "'" & tipo_cai & "', "
    strSQL = strSQL & "'" & num_ven & "', "
    strSQL = strSQL & "'" & Esp_cai & "', "
    strSQL = strSQL & "'" & cod_con & "',"
    strSQL = strSQL & F_GravaDec(entrada_cai) & ","
    strSQL = strSQL & F_GravaDec(troco_cai) & ","
    strSQL = strSQL & "'" & strIP_Usu & "', " 'ip da mákina
    strSQL = strSQL & "'" & caixa & "', " '
    strSQL = strSQL & "'" & turno & "', " '
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & F_GravaHora(Time) & ")"
    
    
    banco.Execute strSQL
    F_GravaCai = sSeq
    Exit Function
trata_erro:
    
End Function

Sub P_BaixaVenda(num_ven As String)
    Dim strSQL As String
    Dim rpesq As Recordset
    Dim sEst As String
    Dim cEst As Currency
    strSQL = "select * from tab_itvendas where num_ven = '" & num_ven & "' "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
            
    With rpesq
        Do While Not .EOF
            If !cod_pro <> "" Then
                sEst = F_BuscaUmCampo("er_pro", "tab_produtos", "cod_pro='" & !cod_pro & "'")
                If sEst = "" Then sEst = "0"
                cEst = CCur(sEst) - !qtd_ite
                If cEst < 0 Then
                    cEst = 0
                End If
            End If
            
            
            
            
            .MoveNext
        
            
        Loop
        .Close
    End With
        
    
        
        strSQL = "update tab_vendas set "
        
        strSQL = strSQL & " situacao_ven='4', "
        strSQL = strSQL & " dataefet_ven=" & F_GravaData(Date)
        strSQL = strSQL & " where num_ven='" & num_ven & "' "
        banco.Execute strSQL
    
End Sub
Sub P_BaixaSaida(seq_sai As String)
    Dim strSQL As String
    Dim rpesq As Recordset
    On Error GoTo trata_erro
    strSQL = "select * from tab_itvendas where num_ven='" & seq_sai & "' "
    
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        If .RecordCount > 0 Then
            Do While Not .EOF
                strSQL = "update tab_produtos set ef_pro=ef_pro-" & str(!qtd_ite)
            
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            banco.Execute strSQL
                .MoveNext
            Loop
            .Close
            Exit Sub
        End If
    End With
    
    strSQL = "select * from tab_itsainf where seq_sai = '" & seq_sai & "' "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
            
    With rpesq
        Do While Not .EOF
            strSQL = "update tab_produtos set ef_pro=ef_pro-" & str(!qtd_its)
            
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            banco.Execute strSQL
            
            
            
            
            .MoveNext
        
            
        Loop
        .Close
    End With
        
    Exit Sub
trata_erro:
        
        
    
End Sub



Sub P_InfPro(cod_pro As String)
    On Error Resume Next
        
        If cod_pro = "" Then Exit Sub
        frmInfPro.Tag = cod_pro
        frmInfPro.caption = "Informações sobre: " & F_BuscaUmCampo("NOme_Pro", "tab_produtos", _
            "cod_pro='" & cod_pro & "'")
        
        frmInfPro.Show 1
End Sub
Function F_EmiteCupomVenda(sNum_ven As String, cPercIcms As Currency, Optional strNome_Cli As String = "", _
Optional strCNPJ As String = "", Optional crcValorPago As Currency = 0, Optional bVendaEsp As Boolean = False, Optional bTEF As Boolean = False) As Boolean
    'Esta função emite o cupom fiscal da nota
    Dim cValorPago As Currency
    Dim blnResp As Boolean
    Dim intCont As Integer, strST As String, strInf As String
    Dim strCod As String, strDescr As String, strUnidade As String
    Dim crcQuantidade As Currency, crcUnitario As Currency
    Dim strFormaP As String
    Dim crcDesconto As Currency
    Dim strCPF As String
    Dim sMsgECF As String
    Dim sTipoMsg As String
    Dim strSQL As String
    Dim rpesq As Recordset
    Dim cEsp As Currency
    Dim i As Integer
    Dim sTipoQtd As String
    crcCpm = 0
    strNum_Cpm = Space(6)
    i = ecf_NumeroCupom(strNum_Cpm)
    If i = 0 Then
        MsgBox "Não foi possível encontrar o numero do cupom!", vbExclamation
        F_EmiteCupomVenda = False
        P_Status ""
        Exit Function
    End If
    If F_BuscaUmCampo("seq_cup", "tab_cupons", "num_ven='" & _
        sNum_ven & "'") <> "" Then
        MsgBox "Este cupom já foi impresso!", vbExclamation
        Exit Function
    End If
    strSQL = "select   tv.*,ti.unit_ite,ti.qtd_ite,ti.seq_ite,(ti.unit_ite*ti.qtd_ite) as total, "
    strSQL = strSQL & "tp.* from   tab_produtos as tp, tab_itvendas as ti, tab_vendas as tv "
    strSQL = strSQL & " where  ti.cod_pro=tp.cod_pro and ti.num_ven='" & sNum_ven & "' "
    strSQL = strSQL & " and tv.num_ven=ti.num_ven "
    
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If rpesq.RecordCount = 0 Then
        MsgBox "Não existem itens nesta venda!", vbExclamation
        rpesq.Close
        Exit Function
    End If
    sTipoMsg = F_BuscaUmCampo("tipomsgecf_par", "tab_Parametros", "")
    
    strInf = strInf & F_Centra(strNome_Cli, _
    48, "-") & Chr(13) & Chr(10)
    If Len(strCNPJ) = 18 Then
        strCPF = "CNPJ: " & strCNPJ
    ElseIf Len(strCNPJ) = 14 Then
        strCPF = "CPF: " & strCNPJ
    End If
    
    strInf = strInf & Chr(13) & F_Centra(strCPF, 48, "")
    
    If sTipoMsg = "1" Then
        strInf = strInf & F_Centra("Tenha " & _
        IIf(Left(F_Cumprimenta(), 3) = "Bom", _
        "um ", "uma ") & F_Cumprimenta(), 48, "-") & Chr(13) & Chr(10)
    Else
        sMsgECF = F_BuscaUmCampo("msgecf_Par", "tab_parametros", "")
        strInf = strInf & sMsgECF
    End If
    strInf = strInf & Chr(13) + Chr(10) & "Caixa: " & strCod_Cai & " Atendente: " & strNome_Usu
    'Carrega o form de formas de pagamento
    'Passando o valor total da nota
    'frmFormas_Pagamento.lblRestante = rPesq!valor_ven
    'frmFormas_Pagamento.txtbxValor.Text = rPesq!valor_ven
    'frmFormas_Pagamento.Show 1
    'strFormaP = frmFormas_Pagamento.Tag
    
    'Unload frmFormas_Pagamento
    If crcValorPago = 0 Then
        cValorPago = rpesq!valor_ven
    Else
        cValorPago = crcValorPago
    End If
    
    If bVendaEsp = True Then
        With rpesq
            cValorPago = 0
            Do While Not .EOF
                cValorPago = cValorPago + (!cf_pro * !qtd_ite) + (!cf_pro * !qtd_ite * 0.3)
                
                .MoveNext
            Loop
            cValorPago = Round(cValorPago, 2)
            .MoveFirst
        End With
        crcCpm = cValorPago
    End If
         
    
    P_Status "Emitindo Cupom Fiscal"

    i = ecf_AbreCupom(strCPF)
    If i <> 1 Then
        MsgBox "Erro na abertura do cupom!", vbExclamation
        F_EmiteCupomVenda = False
        Exit Function
    End If
    
    

    With rpesq
            Do While Not .EOF
            
            strCod = !cod_pro
            
            strDescr = Left(!nome_pro, 50)
            
            
            If !st_pro = "R" Or !st_pro = "N" Or !st_pro = "F" Then
                strST = "FF"
            ElseIf !st_pro = "I" Then
                strST = "II"
            ElseIf !st_pro = "" Then
                strST = "01"
            End If
            If Int(!qtd_ite) <> CCur(!qtd_ite) Then
                sTipoQtd = "F"
            Else
                sTipoQtd = "I"
            End If
            
            If !st_pro = "T" Then
                
                strST = F_IcmsProd(!cod_pro, 3)
            ElseIf !st_pro = "R" Then
                strST = "FF" 'chamada para produto retido. (subst trib.)
            ElseIf !st_pro = "I" Then
                strST = "II" 'chamada para produto isento
            End If
            
            strUnidade = !Unidade_pro
            
            crcQuantidade = !qtd_ite
            
            
            If bVendaEsp = False Then
                crcUnitario = !unit_ite
            Else
                crcUnitario = Round(!cf_pro + (!cf_pro * 0.3), 2)
            End If
            
            crcDesconto = Round(!total / !valor_ven * !desconto_ven, 2)
            
            'blnResp = F_VendaItem(strCod, strDescr, strST, _
            crcQuantidade, crcUnitario, strUnidade, crcDesconto)
            i = ecf_VendeItem(strCod, strDescr, strST, sTipoQtd, _
             IIf(sTipoQtd = "F", Format(crcQuantidade, "###,##0.000"), CStr(crcQuantidade)), 3, Format(crcUnitario, "###,##0.000"), "%", "0")
            
            If i = 0 Then
                MsgBox "Erro de comunicação.", vbCritical
                Bematech_FI_CancelaCupom
                Exit Function
            ElseIf i = -2 Then
                MsgBox "Parâmetro inválido na função.!", vbCritical
                Bematech_FI_CancelaCupom
                Exit Function
            ElseIf i = -3 Then
                MsgBox "Alíquota não programada."
                Bematech_FI_CancelaCupom
                Exit Function
            ElseIf i = -4 Then
                MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório de sistema do Windows."
                Bematech_FI_CancelaCupom
                Exit Function
            ElseIf i = -5 Then
                MsgBox "Erro ao abrir a porta de comunicação."
                Bematech_FI_CancelaCupom
                Exit Function
            ElseIf i = -27 Then
                MsgBox "Status da impressora diferente de 6,0,0 (ACK, ST1 e ST2)."
                Bematech_FI_CancelaCupom
                Exit Function
            End If
             
            
            .MoveNext
        Loop
    End With
    If bTEF Then
        F_EmiteCupomVenda = True
        Exit Function
    End If
    
    i = ecf_IniciaFechamentoCupom("D", "$", CStr(crcDesconto))
    If i <> 1 Then
          MsgBox "Erro em ecf_iniciaFechamentoCupom, mod: = BancoDados Função F_EmiteCupomVenda", vbCritical
    End If
      
      
    i = ecf_EfetuaFormaPagamento("Dinheiro", F_SoNumero(FormatVal(cValorPago)))
    If i <> 1 Then
          MsgBox "Erro em ecf_efetuaFormaPagamento, mod: = BancoDados Função F_EmiteCupomVenda", vbCritical
    End If
    i = ecf_TerminaFechamentoCupom(strInf)
    If i <> 1 Then
          Select Case i
        Case 0: MsgBox "0: Parâmetro inválido na função!", vbCritical
        Case -2: MsgBox "-2: O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório de sistema do Windows.", vbCritical
        Case -4: MsgBox "-4 Erro de comunicação!", vbCritical
        Case -5: MsgBox "-5: Erro ao abrir a porta de comunicação!", vbCritical
        Case -27: MsgBox "-27: Status da impressora diferente de 6,0,0 (ACK, ST1 e ST2).", vbCritical
        Case -36: MsgBox "-36: Forma de pagamento não finalizada!", vbCritical
      End Select
    End If
    

    F_EmiteCupomVenda = True
End Function

Function F_GravaSaidaOP(cod_cli As String, num_sai As String, tipo_sai As String, valor_sai As String, _
    desconto_sai As String, icms_sai As String, num_ven As String, cod_usu As String, _
    data_sai As String, hora_sai As String, Cod_Cai As String, cod_tur As String) As String
    On Error GoTo trata_erro
    strSQL = "inser into tab_saidasop (cod_cli, num_sai, tipo_sai, valor_sai, desconto_sai, icms_sai, "
    strSQL = strSQL & " num_ven, cod_usu, data_sai, cod_cai, hora_sai, cod_tur, seq_sai)values("
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & "'" & num_sai & "',"
    strSQL = strSQL & "'" & tipo_sai & "',"
    strSQL = strSQL & F_GravaDec(valor_sai) & ","
    strSQL = strSQL & F_GravaDec(desconto_sai) & ","
    strSQL = strSQL & F_GravaDec(icms_sai) & ","
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & "'" & cod_usu & "',"
    strSQL = strSQL & F_GravaData(data_sai) & ","
    strSQL = strSQL & "'" & hora_sai & "',"
    strSQL = strSQL & "'" & Cod_Cai & "',"
    strSQL = strSQL & "'" & cod_tur & "',"
    strSQL = strSQL & "'')"
    banco.Execute strSQL
    
    F_GravaSaidaOP = F_BuscaUmCampo("last_insert_id()", "tab_saidasop", "")
    Exit Function
trata_erro:
    F_GravaSaidaOP = ""

End Function

Function F_GravaItSaiOP(seq_sai As String, num_ven As String, cod_pro As String, _
    qtd_its As String, unit_its As String, icms_its As String) As Boolean

    On Error GoTo trata_erro
    strSQL = "inser into tab_saidasop (seq_its, seq_sai,cod_pro, "
    strSQL = strSQL & " qtd_its, unit_its, icms_its)values("
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'" & seq_sai & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(qtd_its) & ","
    strSQL = strSQL & F_GravaDec(unit_its) & ","
    strSQL = strSQL & F_GravaDec(icms_its) & ")"
    
    banco.Execute strSQL
    
    F_GravaItSaiOP = True
    Exit Function
trata_erro:
    F_GravaItSaiOP = False

End Function
Sub P_GravaSaida(num_ven As String, num_sai As String, Optional tipo_sai As String = "C")
    Dim strSQL As String
    Dim rpesq As Recordset
    Dim sSeq As String
    
    strSQL = "select * from tab_vendas where num_ven='" & num_ven & "' "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rpesq
        If .RecordCount > 0 Then
            sSeq = F_GravaSaidaOP(!cod_cli, num_sai, tipo_sai, !valor_ven, !desconto_ven, 0, num_ven, strCod_Usu, _
            Date, Format(Time, "hh:mm"), strCod_Cai, strCod_Tur)
            If sSeq = "" Then
                MsgBox "Erro na gravação da saída!", vbExclamation
                Exit Sub
            End If
            
        End If
        .Close
    End With
    strSQL = "select * from tab_itvendas where num_ven='" & num_ven & "' "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        Do While Not .EOF
            If Not F_GravaItSaiOP(sSeq, num_ven, !cod_pro, !qtd_its, !unit_its, !icms_ite) Then
            End If
            .MoveNext
        Loop
        .Close
    End With
    
End Sub

Sub P_GravaSaidaAntigo(num_ven As String, nota_sai As String, Optional bVendaEsp As Boolean = False)
    Dim strSQL As String
    Dim sCod As String
    Dim sUnit As String
    Dim sDesc As String
    Dim sQtd As String
    Dim iCount As Integer
    Dim rpesq As Recordset
    Dim cValor As Currency
    Dim sSeq As String
    'On Error GoTo trata_erro
    On Error Resume Next
    strSQL = "select * from tab_vendas where num_ven='" & num_ven & "' "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    nota_sai = "      "
    iCount = ecf_NumeroCupom(nota_sai)
    iCount = 0
    With rpesq
    
        strSQL = "Insert Into TAB_saidas (seq_sai,"
        strSQL = strSQL & "COD_CLI, NOTA_SAI, SERIE_SAI, EMISSAO_SAI, "
        strSQL = strSQL & "VALOR_SAI,"
        strSQL = strSQL & "DESCONTO_SAI, "
        strSQL = strSQL & "NUM_VEN,COD_USU, "
        strSQL = strSQL & "cod_cai,cod_tur,hora_sai) Values ( "
        sSeq = ChaveSQL("tab_saidas", "seq_sai", "")
        strSQL = strSQL & sSeq & ","
        strSQL = strSQL & "'" & !cod_cli & "', "
        strSQL = strSQL & "'" & Left(nota_sai, 6) & "', "
        strSQL = strSQL & "'CUPM', "
        strSQL = strSQL & F_GravaData(Date) & ", "
        If bVendaEsp = True Then
            strSQL = strSQL & F_GravaDec(crcCpm) & ", "
            strSQL = strSQL & F_GravaDec(0#) & ", "
        Else
            strSQL = strSQL & F_GravaDec(!valor_ven) & ", "
            strSQL = strSQL & F_GravaDec(!desconto_ven) & ", "
        End If
        strSQL = strSQL & "'" & num_ven & "', "
        
        strSQL = strSQL & "'" & strCod_Usu & "', "
        strSQL = strSQL & "'" & strCod_Cai & "', "
        strSQL = strSQL & "'" & strCod_Tur & "', "
        strSQL = strSQL & F_GravaDataTempo(Now) & ")"
    
        
    
        banco.Execute strSQL
        .Close
    End With
    strSQL = "select * from tab_itvendas where num_ven='" & num_ven & "' order by seq_ite "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rpesq
        Do While Not .EOF
            iCount = iCount + 1 'sequencial
            strSQL = "insert into tab_itsai(seq_sai,cod_pro,nota_sai,serie_sai,unit_its,"
            strSQL = strSQL & "qtd_its,seq_its,cod_cai,icms_its)values("
            strSQL = strSQL & sSeq & ","
            strSQL = strSQL & "'" & !cod_pro & "',"
            strSQL = strSQL & "'" & Left(nota_sai, 6) & "',"
            strSQL = strSQL & "'CUPM',"
            If frmTroco.pc.BackColor = vbBlue Then
                strSQL = strSQL & F_GravaDec(CCur(F_BuscaUmCampo("c1_pro", "tab_composicao", _
                "cod_pro ='" & !cod_pro & "'")) * 1.3) & ","
            Else
                strSQL = strSQL & F_GravaDec(!unit_ite) & ","
            End If
            strSQL = strSQL & F_GravaDec(!qtd_ite) & ","
            strSQL = strSQL & iCount & ","
            strSQL = strSQL & "'" & strCod_Cai & "', "
            strSQL = strSQL & F_GravaDec(F_IcmsProd(!cod_pro, 3)) & ")"
            
            banco.Execute strSQL
            
            strSQL = "update tab_produtos set ef_pro=ef_pro-" & F_GravaDec(!qtd_ite)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            banco.Execute strSQL
            'baixar também a saída
            .MoveNext
        Loop
        .Close
    End With
Exit Sub
trata_erro:
    On Error Resume Next
    Open "C:\erro.txt" For Output As #2
        Print #2, Err.Description & Chr(13) + Chr(10)
        Print #2, strSQL
    Close #2
    MsgBox "Ocorreu um erro na gravação, e foi gerado o arquivo c:\erro.txt", vbExclamation
    Exit Sub
End Sub

Function F_AjustaSaldoCaixa(Cod_Cai As String, sData) As Currency
    Dim strSQL As String
    Dim rpesq As Recordset
    Dim cSaldo As Currency
    strSQL = "Select max(seq_cai) as max  from tab_caixa where data_cai <=" & F_GravaData(sData)
    strSQL = strSQL & "  and cod_cai='" & Cod_Cai & "' "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        If .RecordCount = 0 Then
            strSQL = "Select min(seq_cai) as max  from tab_caixa "
            strSQL = strSQL & "  where cod_cai='" & Cod_Cai & "' "
            Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
            strSQL = F_BuscaUmCampo("valor_cai", "tab_caixa", "seq_cai='" & _
                rpesq!max & "'")
        Else
            strSQL = F_BuscaUmCampo("saldo_cai", "tab_caixa", "seq_cai='" & _
                rpesq!max & "'")
            If strSQL = "" Then
                strSQL = "0"
            End If
            cSaldo = CCur(strSQL)
        
        
        End If
            
    End With
    strSQL = "select * from tab_caixa  where data_cai>=" & F_GravaData(sData)
    strSQL = strSQL & " and cod_Cai='" & Cod_Cai & "' order by data_cai,seq_cai "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        Do While Not .EOF
            If !tipo_cai = "C" Then
                cSaldo = cSaldo + !valor_Cai
            Else
                cSaldo = cSaldo - !valor_Cai
            End If
            strSQL = "update tab_caixa set saldo_cai=" & F_GravaDec(cSaldo)
            strSQL = strSQL & " where seq_Cai='" & !seq_cai & "'"
            banco.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
End Function
Function F_GravaFinanceiro(tipo_fin As String, data_fin As String, _
    valor_fin As String, historico_fin As String, _
    Esp_fin As String, cod_con As String, seq_lan As String, _
    seq_che As String, num_ven As String, cod_pla As String) As String
    
    Dim strSQL As String
    Dim sSeq  As String
    If Esp_fin = "1" Then
        Esp_fin = "DIN"
    ElseIf Esp_fin = "2" Then
        Esp_fin = "CHQ"
    End If
    If tipo_fin = "R" Or tipo_cai = "+" Then
        tipo_fin = "C"
    End If
    If tipo_fin = "P" Or tipo_cai = "-" Then
        tipo_fin = "D"
    End If
    If cod_pla = "" Then
        If tipo_fin = "D" Then
            cod_pla = strPlaNaoClassD
        Else
            cod_pla = strPlaNaoClassC
        End If
    End If
    strSQL = "insert into tab_financeiro(seq_fin,data_fin,valor_fin,saldo_fin,historico_fin,tipo_fin,"
    strSQL = strSQL & "esp_fin,seq_lan,NUM_ven,cod_con,cod_pla)Values("
    sSeq = ChaveSQL("tab_financeiro", "seq_fin", "000000#")
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(data_fin) & ","
    strSQL = strSQL & F_GravaDec(valor_fin) & ","
    If tipo_fin = "C" Then
        strSQL = strSQL & " saldo_fIN + " & F_GravaDec(valor_fin) & ","
    Else
        strSQL = strSQL & " saldo_FIN + " & F_GravaDec(valor_fin) & ","
    End If
    strSQL = strSQL & "'" & historico_fin & "',"
    strSQL = strSQL & "'" & tipo_fin & "', "
    
    strSQL = strSQL & "'" & Esp_fin & "', "
    strSQL = strSQL & "'" & seq_lan & "', "
    strSQL = strSQL & "'" & seq_ven & "', "
    strSQL = strSQL & "'" & cod_con & "',"
    strSQL = strSQL & "'" & cod_pla & "') "
    
    F_GravaFinanceiro = sSeq
    banco.Execute strSQL
    
End Function


Function F_CalculaSaldo(data_fin As String, cod_con As String) As Currency
    Dim strSQL As String
    Dim rpesq As Recordset
    Dim cSaldo As Currency
    Dim bSaldo As Boolean
    On Error Resume Next
    strSQL = "select Max(data_fin) as max from tab_financeiro where data_fin < "
    strSQL = strSQL & F_GravaData(data_fin) & " and cod_Con='" & cod_con & "' "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If IsNull(rpesq!max) Then
        strSQL = "select Min(data_fin) as max from tab_financeiro  where cod_Con='" & cod_con & "' "
        Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
        strSQL = "select min(seq_fin) as max from tab_financeiro where data_fin = "
        strSQL = strSQL & F_GravaData(rpesq!max) & " and cod_Con='" & cod_con & "' "
        Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
        bSaldo = False
    Else
        bSaldo = True
        strSQL = "select max(seq_fin) as max from tab_financeiro where data_fin = "
        strSQL = strSQL & F_GravaData(rpesq!max) & " and cod_Con='" & cod_con & "' "
        Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
'        strSQL = "select max(Seq_fin) as max from tab_financeiro where data_fin = "
'        strSQL = strSQL & F_GravaData(rPesq!Max)
'        Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    End If
        strSQL = "select * from tab_financeiro where Seq_fin = '" & rpesq!max & "'" & " and cod_Con='" & cod_con & "' "
        Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
        cSaldo = rpesq!saldo_fin
        strSQL = "select * from tab_financeiro where data_fin >= " & F_GravaData(data_fin)
        strSQL = strSQL & " and cod_Con='" & cod_con & "' "
        strSQL = strSQL & " Order by  data_fin, Seq_fin"
        Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
        
        With rpesq
            If bSaldo = False Then
                cSaldo = 0
            End If
            Do While Not .EOF
                If !tipo_fin = "C" Then
                    cSaldo = cSaldo + !valor_fin
                Else
                    cSaldo = cSaldo - !valor_fin
                End If
                strSQL = "update tab_financeiro set saldo_fin = " & F_GravaDec(cSaldo)
                strSQL = strSQL & " where Seq_fin = '" & !seq_fin & "'"
                banco.Execute strSQL
            .MoveNext
            Loop
            .Close
        End With
      
    F_CalculaSaldo = cSaldo
    
    
End Function

Sub P_MediaDiaria(sDataIni As String, sDataFin As String, Optional sMar As String, _
    Optional sCat As String, Optional sPro As String, Optional sFor As String = "")
                    
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim iDias As Integer
    Dim rpesq As Recordset
    Dim sDias_Trab As String
    P_Status "Buscando produtos movimentados no período..."
    strSQL = "update tab_produtos set mmd_pro = 0"
    banco.Execute strSQL
    sDias_Trab = F_BuscaUmCampo(" count(distinct dataefet_ven)", "tab_vendas", _
                " dataefet_ven >=" & F_PesqData(sDataIni) & " and dataefet_ven <=" & F_PesqData(sDataFin))
    strSQL = " select"
    strSQL = strSQL & "  ti.cod_pro,"
    strSQL = strSQL & " tp.cod_mar,"
    strSQL = strSQL & " sum(ti.Qtd_ite) as qtd,"
    strSQL = strSQL & " Count(tv.dataefet_ven) as dias "
    strSQL = strSQL & " From "
    
    strSQL = strSQL & " tab_vendas as tv,"
    strSQL = strSQL & " tab_itvendas As ti,tab_produtos as tp "
    strSQL = strSQL & " Where ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and ti.num_ven = tv.num_ven "
    strSQL = strSQL & " and "
    strSQL = strSQL & " tv.dataefet_ven >=" & F_PesqData(sDataIni)
    strSQL = strSQL & " and tv.dataefet_ven <=" & F_PesqData(sDataFin)
'    If sFor <> "" Then
'        strSQL = strSQL & " and tic.seq_not=tc.seq_not "
'        strSQL = strSQL & " and tic.cod_pro=ti.cod_pro "
'        strSQL = strSQL & " and tc.cod_for='" & sFor & "'"
'    End If
    If sCat <> "" Then
        strSQL = strSQL & " and tp.cod_cat='" & sCat & "' "
    End If
    
    If sMarca <> "" Then
        strSQL = strSQL & " and tp.cod_mar='" & sMar & "' "
    End If
    If sNome <> "" Then
        strSQL = strSQL & " and tp.nome_pro like '" & sPro & "%' "
    End If
    strSQL = strSQL & " Group By"
    strSQL = strSQL & " ti.cod_pro order by ti.cod_pro"

    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        Do While Not .EOF
            strSQL = "update tab_produtos set mmd_pro = " & str(Round(!qtd / sDias_Trab, 4)) & ","
            strSQL = strSQL & "ipm_pro=" & F_GravaData(sDataIni) & ","
            strSQL = strSQL & "fpm_pro=" & F_GravaData(sDataFin)
            strSQL = strSQL & " where cod_pro = '" & !cod_pro & "' "
            
            banco.Execute strSQL

            If .AbsolutePosition Mod 50 = 0 Or .AbsolutePosition = .RecordCount - 1 Then
                With frmMediaConsumo_Mer
                    .Image1.Width = .Label3.Width * ((rpesq.AbsolutePosition + 1) / rpesq.RecordCount)
                    .Image1.Refresh
                    .Label3.caption = Round((rpesq.AbsolutePosition + 1) * 100 / rpesq.RecordCount) & " %"
                End With
            End If
            .MoveNext
        
        Loop
        frmMediaConsumo_Mer.Label3.caption = .RecordCount & " Registros"
        .Close
    P_Status
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
    P_Status ""
    
End Sub
Function F_CalculaJuros(vencimento As String, data As String, valor As Currency, _
    Optional bJuroMaiorQueZero As Boolean = True) As Currency
    Dim sTxJuros As String
    Dim iDias As Integer
    On Error GoTo trata_erro
    sTxJuros = F_BuscaUmCampo("juros_par", "tab_parametros", "")
    If sTxJuros = "" Then
        sTxJuros = "0"
    End If
    iDias = DateDiff("D", vencimento, Date)
    F_CalculaJuros = (iDias / 30) * CCur(sTxJuros) / 100 * valor
    If F_CalculaJuros < 0 And bJuroMaiorQueZero = True Then
        F_CalculaJuros = 0
    End If
    Exit Function
trata_erro:
    F_CalculaJuros = 0
End Function

Sub P_Class(valor_cla As Currency, sTipo As String, seq As String)
    'On Error Resume Next
    
    If sTipo = "R" Then
        sTipo = "C"
    End If
    If sTipo = "P" Then
        sTipo = "D"
    End If
    frmClass.cmndbtnConfirma.Tag = sTipo
    frmClass.lblTOtal.caption = FormatVal(valor_cla)
    frmClass.Tag = seq
    frmClass.Show 1

    If frmClass.Tag <> "" Then
        MsgBox "Lançamento Classificado!", vbInformation
        
    Else
        MsgBox "O Lançamento não foi classificado!", vbInformation
    End If
    
    Unload frmClass
End Sub

Sub P_PreenchePlanoContas(obj As Object, optnCred As Boolean)
    On Error GoTo trata_erro
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select * from tab_planocontas  "
    If optnCred = True Then
        strSQL = strSQL & " where tipo_pla='C' "
    Else
        strSQL = strSQL & " where tipo_pla='D' "
    End If
    strSQL = strSQL & " order by nome_pla"
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    With rpesq
        Do While Not .EOF
            obj.AddItem rpesq!nome_pla
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub

Function F_PlanoConta(optnNome As Boolean, strPla As String, tipo_pla As String) As String
    
    If optnNome Then
        F_PlanoConta = F_BuscaUmCampo("nome_pla", "tab_planocontas", "tipo_pla='" & _
            tipo_pla & "' and cod_pla='" & strPla & "'")
    Else
        F_PlanoConta = F_BuscaUmCampo("cod_pla", "tab_planocontas", "tipo_pla='" & _
            tipo_pla & "' and nome_pla='" & strPla & "'")
    End If
End Function


Function F_GravaDevolucao(num_ven As String, cod_pro As String, qtd_ite As Currency, _
        qtd_dev As Currency, data_dev As Date, gerado_Dev As Currency, tipo_dev As String) As Boolean

    Dim strSQL As String
    Dim rpesq As Recordset
    'On Error GoTo trata_erro
    
    strSQL = "insert into tab_devolucoes(num_ven,cod_pro,qtd_ite,qtd_dev,data_dev, gerado_dev,"
    strSQL = strSQL & "cod_usu)Values("
    
    strSQL = strSQL & "'" & num_ven & "', "
    strSQL = strSQL & "'" & F_SoNumero(cod_pro) & "', "
    strSQL = strSQL & F_GravaDec(qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(qtd_dev) & ","
    
    strSQL = strSQL & F_GravaDataTempo(Now) & ",   "
    
    strSQL = strSQL & F_GravaDec(gerado_Dev) & ",   "
    strSQL = strSQL & "'" & strCod_Usu & "')  "
    banco.Execute strSQL
    
    strSQL = "update tab_produtos set er_pro=er_pro+"
    strSQL = strSQL & F_GravaDec(qtd_dev)
    strSQL = strSQL & " where cod_pro='" & cod_pro & "'"
    banco.Execute strSQL
    
    
    
    
    F_GravaDevolucao = True
    
    Exit Function
trata_erro:
    F_GravaDevolucao = False
End Function


Sub P_imprDevolucao(num_ven As String, obj As Object, Optional bVal As Boolean = True)
    Dim strSQL As String
    Dim rpesq As Recordset
    Dim cTotal As Currency
    
    'On Error GoTo trata_erro
    strSQL = "select tv.*,td.qtd_ite,td.tipo_dev,td.qtd_dev,td.data_dev,td.gerado_dev,"
    strSQL = strSQL & "tf.avista_ven,tf.aprazo_ven,tf.finan_ven,tf.valorfinal_ven,"
    strSQL = strSQL & "tp.cod_pro , "
    strSQL = strSQL & "tp.nome_pro from tab_produtos as tp, tab_devolucoes as td,"
    strSQL = strSQL & "tab_vendas as tv,tab_venfec as tf "
    strSQL = strSQL & " where tv.num_ven=tf.num_ven and tv.num_ven=td.num_ven and "
    strSQL = strSQL & " td.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and tv.num_ven = '" & num_ven & "'"
    strSQL = strSQL & " and left(td.data_dev,10) = " & F_GravaData(Date) & ""
    strSQL = strSQL & " order by nome_Pro "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If rpesq.RecordCount = 0 Then
        MsgBox "Não há itens devolvidos nesta venda!", vbExclamation
        rpesq.Close
        Exit Sub
    End If
    With rpesq
        If bVal Then
            P_CabRelTimbr "GUIA DE DEVOLUÇÃO - VENDA " & num_ven, obj
        Else
            P_CabRelTimbr "GUIA DE RECOLHIMENTO DE MERCADORIA - Nº " & num_ven, obj
        End If
        P_ImprLinha obj, "_"
        obj.fontsize = 12
        P_ImprX obj, "Cliente: " & F_BuscaUmCampo("nome_cli", "tab_clientes", _
            "cod_cli='" & !cod_cli & "'"), 10, True
        If bVal Then
            P_ImprX obj, "Venda: " & !num_ven, 120, True
            P_ImprX obj, "Valor: " & FormatVal(!valorfinal_ven), 160, False
        End If
        obj.fontsize = 10
        P_ImprLinha obj, "-"
        P_ImprX obj, "Código", 5
        P_ImprX obj, "Produto", 25
        P_ImprX obj, "T", 135
        P_ImprX obj, "Vendidos", 155, True, True
        P_ImprX obj, "Devolvidos", 180, True, True
        If bVal = True Then
            P_ImprX obj, "Vlr Gerado", 200, False, True
        Else
            P_ImprX obj, "------", 200, False, True
        End If
        P_ImprLinha obj, "-"
        Do While Not .EOF
            P_ImprX obj, Format(!cod_pro, "00#\.###\-#"), 5
            P_ImprX obj, !nome_pro, 25
            P_ImprX obj, !tipo_dev, 135
            P_ImprX obj, FormatVal(!qtd_ite), 155, True, True
            P_ImprX obj, FormatVal(!qtd_dev), 180, True, True
            If bVal Then
                P_ImprX obj, FormatVal(!gerado_Dev), 200, False, True
            Else
                P_ImprX obj, "________", 200, False, True
            End If
            cTotal = cTotal + !gerado_Dev
            .MoveNext
        
        Loop
        .Close
        P_ImprLinha obj
        obj.fontsize = 14
        If bVal Then
            P_ImprX obj, "Total da Devolução............: " & FormatVal(cTotal), 200, False, True
        Else
            P_ImprX obj, "Via do Entregador", 200, False, True
        End If
    End With
    
    obj.EndDoc
    Exit Sub
trata_erro:
    Printer.KillDoc
End Sub


Sub P_PreencheBarCode(cod_pro As String, obj As Object)
    Dim strSQL As String
    Dim rpesq As Recordset
    On Error GoTo trata_erro
    strSQL = "select * from tab_codbarras where cod_pro='" & cod_pro & "' order by codbarra_pro"
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    DoEvents
    With rpesq
        Do While Not .EOF
            obj.AddItem !codbarra_pro
            .MoveNext
        Loop
        
        .Close
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_PreencheEmbalagens(cod_pro As String, obj As Object, Optional bUN As Boolean = False)
    Dim strSQL As String
    Dim rpesq As Recordset
    On Error GoTo trata_erro
    strSQL = "select * from tab_embalagens where cod_pro='" & cod_pro & "' order by qtd_emb"
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    obj.Clear
    DoEvents
    
    With rpesq
        Do While Not .EOF
            obj.AddItem !qtd_emb
            .MoveNext
        Loop
        
        .Close
    If bUN Then
        obj.AddItem "1"
    End If
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Function F_IcmsProd(cod_pro As String, Retorno As TipoRetornoIcms) As String
    Dim strSQL As String
    Dim rpesq As Recordset
    On Error GoTo trata_erro:
    strSQL = "select ti.perc_imp,desc_imp,te.* from tab_produtos as tp, tab_impostos as ti,tab_ecfimp as te"
    strSQL = strSQL & " where tp.cod_imp=ti.cod_imp "
    strSQL = strSQL & " and tp.cod_pro='" & cod_pro & "'"
    strSQL = strSQL & " and ti.cod_imp=te.cod_imp "
    strSQL = strSQL & " and te.impressora_imp='" & strImprFiscal & "'"
    'strSQL = "select * from tab_impostoes where cod_pro='" & cod_pro & "'"
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        If .RecordCount > 0 Then
            If Retorno = 4 Then
                F_IcmsProd = !perc_imp
            ElseIf Retorno = 3 Then
                F_IcmsProd = Format(!perc_imp, "0#") & "00"
            ElseIf Retorno = 2 Then
                F_IcmsProd = !desc_imp
            ElseIf Retorno = 1 Then
                F_IcmsProd = !ecf_imp
            Else
                F_IcmsProd = !cod_imp
            End If
        End If
        .Close
    End With
    Exit Function
trata_erro:
    
 End Function
    Function F_TaxasEntradas() As Currency
    F_TaxasEntradas = crcTxFederativa_Par + crcTxAdm + crcTxMargemSeg_Par
End Function
Function F_CancelaVendaAposCupom(sNum_ven As String, Optional BaixaF As Boolean = False) As Boolean
    Dim iPos As Integer
    Dim strSQL As String
    Dim rpesq As Recordset
    On Error GoTo trata_erro
    'varre toda a venda voltando o estoque real dos itens depois muda para situacao 3
    
    F_CancelaVendaAposCupom = False
    
    
    If F_BuscaUmCampo("status_ven", "tab_vendas", "num_ven='" & sNum_ven & "'") = "3" Then
        MsgBox "Não foi possível cancelar os lançamentos da venda, pois ela já se encontra cancelada!", vbExclamation
        F_CancelaVendaAposCupom = False
        Exit Function
    End If
    
    
   
    
    
    
    strSQL = "select * from tab_itvendas where num_ven='" & sNum_ven & "'  "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        Do While Not .EOF
        
        
            
        
        
        
        
            strSQL = "update tab_produtos set er_pro=er_pro + " & F_GravaDec(CStr(!qtd_ite))
            If BaixaF Then
                strSQL = strSQL & ",ef_pro=ef_pro + " & F_GravaDec(CStr(!qtd_ite))
            End If
        
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            banco.Execute strSQL
        
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "select * from tab_venfec where num_ven='" & sNum_ven & "' "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        If .RecordCount > 1 Then
            MsgBox "Foi encontrado mais que um registro. Os documentos não podem ser deletados", vbExclamation
            Exit Function
        End If
        If .RecordCount > 0 Then
            If !avista_ven > 0 Then
                strSQL = F_GravaCai("D", Date, CStr(!avista_ven), "Venda " & !num_ven & " Cancelada", !num_ven, "DIN", "01", 0, 0, strCod_Cai, strCod_Tur)
            End If
            If !aprazo_ven > 0 Then
                strSQL = "select * from tab_doc1 where num_ven='" & sNum_ven & "' "
                Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
                With rpesq
                    'Do While Not .EOF
                        strSQL = "delete from tab_itdoc1 where seq_doc='" & !seq_doc & "'"
                        banco.Execute strSQL
                        
                        strSQL = "delete from tab_cheques1 where seq_doc='" & !seq_doc & "'"
                        banco.Execute strSQL
                        strSQL = "delete from tab_doc1 where seq_doc='" & !seq_doc & "'"
                        banco.Execute strSQL
                        If banco.RecordsAffected > 2 Then
                            MsgBox "Provavelmente houve um erro. Avise o responsável imediatamente ERRO FINANCEIRO.", vbExclamation
                            
                        End If
                        
                        
                     '   .MoveNext
                    'Loop
                    .Close
                End With
                
            End If
        End If
        .Close
    End With
    strSQL = "update tab_vendas set situacao_ven=3, "
    strSQL = strSQL & " cod_usu='" & strCod_Usu & "' "
    strSQL = strSQL & " where num_ven='" & sNum_ven & "' "
    banco.Execute strSQL '3 = venda cancelada
    F_CancelaVendaAposCupom = True
    Exit Function
trata_erro:
    P_MostraErro
End Function

Sub P_RetornaEstFiscalCancelamento(cupom As String, caixa As String)
    Dim strSQL As String
    
    strSQL = "select * from tab_itsai where nota_sai='" & cupom & "' "
    strSQL = strSQL & " and cod_cai='" & caixa & "' "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rpesq
        Do While Not .EOF
        
            strSQL = "update tab_produtos set ef_pro=ef_pro + " & F_GravaDec(CStr(!qtd_its))
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            banco.Execute strSQL
        
            .MoveNext
        Loop
        .Close
    End With
End Sub
Function F_RetornaLimiteUsoCli(cod_cli As String) As Currency
    Dim strSQL As String
    Dim rpesq As Recordset
    strSQL = "select sum(ti.reman_itd) as total "
    strSQL = strSQL & " from tab_itdoc1 as ti, tab_doc1 as td"
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.cod_ter='" & cod_cli & "' "
    strSQL = strSQL & " and td.tipo_doc='R' "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not IsNull(rpesq!total) Then
        F_RetornaLimiteUsoCli = rpesq!total
    End If
    
    
    strSQL = "select sum(ti.valor_che) as total "
    strSQL = strSQL & " from tab_cheques1 as ti, tab_doc1 as td"
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.cod_ter='" & cod_cli & "' "
    strSQL = strSQL & " and td.tipo_doc='R' "
    strSQL = strSQL & " and ti.dep_che is null "
    Set rpesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not IsNull(rpesq!total) Then
        F_RetornaLimiteUsoCli = rpesq!total + F_RetornaLimiteUsoCli
    End If
    
    
    
    rpesq.Close
End Function


