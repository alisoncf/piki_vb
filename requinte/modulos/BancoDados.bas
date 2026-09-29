Attribute VB_Name = "BancoDados"
Public Const encryptKey = "piki2002"
Option Explicit
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
    
    Conn.Execute strSQL
    F_GeraNovoOrc = sNum_ven
    
    
    
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
    Conn.Execute strSQL
    

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
    Conn.Execute strSQL
    

    F_GeraNovaTransf = sSeq_Tra
    
    
    
End Function

Function F_GeraNovoArranjo() As String
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    sSeq = ChaveSQL("tab_arranjos", "num_arr", "000000#")
    strSQL = "insert into tab_arranjos(num_arr,data_arr,hora_arr,cod_usu,situacao_arr)"
    strSQL = strSQL & "values("
    
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & F_GravaDataTempo(Date) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'1')"
    Conn.Execute strSQL
    

    F_GeraNovoArranjo = sSeq
    Exit Function
trata_Erro:
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
    
    Conn.Execute strSQL
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
    Conn.Execute strSQL
    F_GeraNovoPedido = sNum_ven
    
    
    
End Function

Sub P_GeraLog(sExp As String, sTabela As String)
    Dim strSQL As String
    
    strSQL = "insert into tab_log(seq_log,expressao_log,data_log,tab_log)Values("
    'strSQL = strSQL & ChaveSQL("tab_log", "seq_log", "") & ","
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'" & F_AjustaTexto(sExp) & "',"
    strSQL = strSQL & "" & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & sTabela & "')"
    
    Conn.Execute strSQL
End Sub
Sub P_LogToText()
    Dim fs, a
    Dim strSQL As String
    Dim sNomeArq As String
    Dim rPesq As New ADODB.Recordset
    Dim strChunck As String
    Dim chTest
    P_Status "Gerando arquivo de Log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    sNomeArq = App.path & "\" & "" & Format(Now, "yyyymmddhhmmss") & ".stt"
    Set a = fs.CreateTextFile(sNomeArq, True)
    strSQL = "select * from tab_log where status_log=0"
    abreObjetoAdo rPesq, strSQL
    With rPesq
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
            Conn.Execute strSQL
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
    Dim rPesq As New ADODB.Recordset
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
                Conn.Execute "insert into tab_transferencias values(" & sAux & ")"
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


Function F_BuscaUmCampo(sCampoRetorno As String, sTabela As String, sWhere As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
  
    
    
    On Error GoTo trata_Erro
    strSQL = "Select " & sCampoRetorno & " as Campo from " & sTabela
    If sWhere <> "" Then
        strSQL = strSQL & " Where " & sWhere
        
    End If
    
    
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount < 1 Then
        F_BuscaUmCampo = ""
    Else
        F_BuscaUmCampo = rPesq!campo
    End If
    rPesq.Close
Exit Function
trata_Erro:
        F_BuscaUmCampo = ""
End Function
Function F_BuscaUmCampoOutroBanco(b As database, sCampoRetorno As String, sTabela As String, sWhere As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "Select " & sCampoRetorno & " as Campo from " & sTabela
    If sWhere <> "" Then
        strSQL = strSQL & " Where " & sWhere
    End If
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount < 1 Then
        F_BuscaUmCampoOutroBanco = ""
    Else
        F_BuscaUmCampoOutroBanco = rPesq!campo
    End If
    rPesq.Close
Exit Function
trata_Erro:
        F_BuscaUmCampoOutroBanco = ""
End Function

Function F_GravaDec(sValor) As String
    On Error GoTo trata_Erro
    
    F_GravaDec = str(sValor)
    Exit Function
trata_Erro:
    F_GravaDec = "0"
End Function
Function F_GravaHora(strHora, Optional formato As String = "HH:MM:SS") As String
Dim strDia, strMes, StrAno As String
  strHora = Format(strHora, formato)
  
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
  On Error GoTo trata_Erro
  F_GravaDataTempo = "'" & Format(strData, "yyyy-mm-dd hh:mm:ss") & "' "
  Exit Function
trata_Erro:
  F_GravaDataTempo = "'0000-00-00 00:00:00'"
End Function

Sub P_PreencheLojas(obj As Object, Optional codigo As Boolean = True)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_lojas order by cod_loj "
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    With rPesq
        Do While Not .EOF
            
            If codigo = True Then
                obj.AddItem rPesq!cod_loj
            Else
                obj.AddItem rPesq!cod_loj & " - " & !nome_loj
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_PreencheMarca(obj As Object)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_marcas order by nome_mar "
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    With rPesq
        Do While Not .EOF
            obj.AddItem rPesq!nome_mar
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_PreencheFinanceiras(obj As Object, Optional tipo As String = "")
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_financeiras "
    If tipo <> "" Then
        strSQL = strSQL & " where operacao_fin='" & tipo & "' "
    End If
    strSQL = strSQL & " order by nome_fin "
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    With rPesq
        Do While Not .EOF
            obj.AddItem rPesq!nome_fin
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub


Sub P_PreencheVendedor(obj As Object, Optional opcao As Integer = 0)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset 'se opcao = 0 traz nome completo , se 1 nome de guerra
    strSQL = "select tf.* from tab_funcionarios as tf, tab_cargos as tc "
    strSQL = strSQL & " where tc.cod_car=tf.cod_car and tc.venda_car='S' "
    strSQL = strSQL & " and tf.status_fun='1' "
    If opcao = 0 Then
        strSQL = strSQL & " order by nome_fun "
    Else
        strSQL = strSQL & " order by guerra_fun"
    End If
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    With rPesq
        Do While Not .EOF
            If opcao = 0 Then
                obj.AddItem rPesq!nome_fun
            Else
                obj.AddItem rPesq!guerra_fun
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_PreencheBanco(obj As Object)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_Bancos order by Cod_Bco "
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    With rPesq
        Do While Not .EOF
            obj.AddItem rPesq!cod_bco & "-" & rPesq!nome_bco
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub
Sub P_PreencheTurnos(obj As Object)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_turnos order by Cod_tur "
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    With rPesq
        Do While Not .EOF
            obj.AddItem rPesq!cod_tur & " - " & rPesq!horaini_tur & " às " & rPesq!horafin_tur
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_PreencheCaixas(obj As Object)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select distinct cod_cai,nome_cai from tab_caixas order by Cod_cai "
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    With rPesq
        Do While Not .EOF
            obj.AddItem rPesq!cod_cai & "-" & rPesq!nome_cai
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub
Sub P_PreencheFornecedores(obj As Object)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_fornecedores order by nome_for "
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    With rPesq
        Do While Not .EOF
            obj.AddItem rPesq!nome_for
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_PreencheCargos(obj As Object)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_cargos order by nome_car "
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    With rPesq
        Do While Not .EOF
            obj.AddItem rPesq!nome_car
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_PreencheContas(obj As Object, Optional bFlag As Byte = 1, Optional bCaixa As Boolean = True)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    'Flag: 0 - só o código, 1 - código + nome, 2 - só nome
    'bcaixa:  se true carrega conta 001(caixa),se não oculta
    strSQL = "select * from tab_Contas "
    If bCaixa = False Then
        strSQL = strSQL & " Where cod_con <>  '001'"
    End If
    If bFlag = 0 Or bFlag = 1 Then
        strSQL = strSQL & "order by cod_con "
    
    Else
        strSQL = strSQL & "order by nome_con "
    
    End If
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    With rPesq
        Do While Not .EOF
            If bFlag = 0 Or bFlag = 2 Then
                obj.AddItem rPesq!cod_con
            ElseIf bFlag = 1 Then
                obj.AddItem rPesq!cod_con & "-" & rPesq!nome_con
            Else
                obj.AddItem rPesq!nome_con
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub






Sub P_PreencheFunc(obj As Object, Optional opcao As Integer = 0)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset 'se opcao = 0 traz nome completo , se 1 nome de guerra
    strSQL = "select * from tab_funcionarios "
    If opcao = 0 Then
        strSQL = strSQL & " order by nome_fun "
    Else
        strSQL = strSQL & " order by guerra_fun"
    End If
    strSQL = strSQL & " and status_fun=1 "
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    With rPesq
        Do While Not .EOF
            If opcao = 0 Then
                obj.AddItem rPesq!nome_fun
            Else
                obj.AddItem rPesq!guerra_fun
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub


Function MaxKeySQL(sTab As String, sCampo As String, sWh As String, formato As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select max(" & sCampo & ") as campo  from " & sTab
    If sWh <> "" Then
        strSQL = strSQL & " where " & sWh
    End If
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If IsNull(!campo) Then
            MaxKeySQL = Format(1, formato)
        Else
            
            MaxKeySQL = Format(F_SoNumero(!campo) + 1, formato)
        End If
    End With
    
    
    
    rPesq.Close
    Exit Function
trata_Erro:
    MaxKeySQL = Format(1, formato)
End Function
Function GeraMaxSQL(sCampo As String, Optional formato As String = "") As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select " & sCampo & " as campo  from tab_keys "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        GeraMaxSQL = Format(!campo + 1, formato)
        strSQL = "update tab_keys set " & sCampo & "=" & GeraMaxSQL
        Conn.Execute strSQL
        
        
    End With
    rPesq.Close
End Function
Function ChaveSQL(sTable As String, sField As String, formato As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select value_cod  as campo  from tab_codigos "
    strSQL = strSQL & " where table_cod='" & sTable & "' "
    strSQL = strSQL & " and field_cod='" & sField & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If IsNull(!campo) Or .RecordCount = 0 Then
            strSQL = "insert into tab_codigos(table_cod,field_cod,value_cod)values("
            strSQL = strSQL & "'" & sTable & "',"
            strSQL = strSQL & "'" & sField & "',"
            strSQL = strSQL & "'" & Format("1", formato) & "')  "
            Conn.Execute strSQL
            ChaveSQL = Format("1", formato)
        Else
            
            ChaveSQL = Format(F_SoNumero(!campo) + 1, formato)
            strSQL = "update tab_codigos set value_cod='" & ChaveSQL & "' "
            strSQL = strSQL & " where table_cod='" & sTable & "' "
            strSQL = strSQL & " and field_cod='" & sField & "' "
            Conn.Execute strSQL
        End If
    End With
    rPesq.Close
End Function
Sub FlushCodigos()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "select * from tab_codigos"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            
           
            
            .MoveNext
        Loop
    End With
End Sub
Function F_GravaCHV(seq_cai As String, num_che As String, cod_bco As String, _
    age_che As String, cco_che As String, cpf_che As String, nome_che As String, Valor_che As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
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
    Conn.Execute strSQL
    Exit Function
trata_Erro:
    F_GravaCHV = False
End Function


Function F_GravaCai(tipo_cai As String, data_cai As String, _
    valor_Cai As String, historico_cai As String, num_ven As String, _
    esp_cai As String, cod_con As String, entrada_cai As String, saida_cai As String, _
    CAIXA As String, turno As String, cod_pla As String, visivel_cai As String, Optional desconto_cai As String = "0", Optional hora_cai As String = "") As String
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim sSeq  As String
    Dim entrada As String
    Dim saida As String
    
    saida = saida_cai
    entrada = entrada_cai
    If tipo_cai = "R" Or tipo_cai = "+" Then
        tipo_cai = "C"
    End If
    If tipo_cai = "P" Or tipo_cai = "-" Then
        tipo_cai = "D"
    End If
    If esp_cai = "1" Then
        esp_cai = "DIN"
    End If
    If hora_cai = "" Then
        hora_cai = Now
    End If
    
    If tipo_cai = "D" Then
        saida = valor_Cai
        
    Else
        entrada = valor_Cai
        If esp_cai = "DEP" Or esp_cai = "DUP" Or esp_cai = "RCC" Or esp_cai = "RCD" Or esp_cai = "CRP" Then
            saida = entrada
        End If
    End If
    
    
    
    
    
    strSQL = "insert into tab_caixa(seq_cai,data_cai,valor_Cai,historico_Cai,tipo_Cai,"
    strSQL = strSQL & "num_ven,esp_cai,cod_con,entrada_Cai,saida_cai,ip_cai,cod_cai,cod_tur,cod_pla,visivel_cai,cod_usu,desconto_cai,hora_cai)Values("
    sSeq = ChaveSQL(TAB_CAIXA, "seq_cai", "000000#")
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(data_cai) & ","
    strSQL = strSQL & F_GravaDec(valor_Cai) & ","
    strSQL = strSQL & "'" & Left(historico_cai, 50) & "',"
    strSQL = strSQL & "'" & tipo_cai & "', "
    strSQL = strSQL & "'" & num_ven & "', "
    strSQL = strSQL & "'" & esp_cai & "', "
    strSQL = strSQL & "'" & cod_con & "',"
    strSQL = strSQL & F_GravaDec(entrada) & ","
    strSQL = strSQL & F_GravaDec(saida) & ","
   
    strSQL = strSQL & "'" & "', "  'ip da mákina
    strSQL = strSQL & "'" & CAIXA & "', " '
    strSQL = strSQL & "'" & turno & "', " '
    strSQL = strSQL & "'" & cod_pla & "', "  '
    strSQL = strSQL & "'" & visivel_cai & "', "  '
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & F_GravaDec(desconto_cai) & ","
    strSQL = strSQL & F_GravaDataTempo(hora_cai) & ")"
    
    Conn.Execute strSQL
    F_GravaCai = sSeq
    Exit Function
trata_Erro:
    
End Function

Sub P_BaixaVenda(num_ven As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sEst As String
    Dim cEst As Currency
    strSQL = "select cod_pro,qtd_ite from tab_itvendas where num_ven = '" & num_ven & "' "
    abreObjetoAdo rPesq, strSQL
            
    With rPesq
        Do While Not .EOF
            
            
            strSQL = "update tab_produtos set er_pro=er_pro-" & F_GravaDec(!qtd_ite)
            
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Conn.Execute strSQL
            
            
            
            
            .MoveNext
        
            
        Loop
    End With
        
    
        
        
    
End Sub
Sub P_BaixaSaida(num_ven As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select * from tab_itvendas where num_ven = '" & num_ven & "' "
    abreObjetoAdo rPesq, strSQL
            
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_composicao set e1_pro=e1_pro-" & str(!qtd_ite)
            
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            'conn.execute strSQL
            
            
            
            
            .MoveNext
        
            
        Loop
    End With
        
    
        
        
    
End Sub



Sub P_InfPro(cod_pro As String)
    On Error Resume Next
        
        If cod_pro = "" Then Exit Sub
        frmInfPro.tag = cod_pro
        frmInfPro.caption = "Informações sobre: " & F_BuscaUmCampo("NOme_Pro", "tab_produtos", _
            "cod_pro='" & cod_pro & "'")
        
        frmInfPro.Show 1
End Sub
Function F_EmiteCupomVenda(sNum_ven As String, cPercIcms As Currency, Optional strNome_Cli As String = "", _
Optional strCNPJ As String = "", Optional crcValorPago As Currency = 0, Optional bVendaEsp As Boolean = False, _
    Optional bTef As Boolean = False) As Boolean
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
    Dim rPesq As New ADODB.Recordset
    Dim cEsp As Currency
    Dim i As Integer
    Dim sTipoQtd As String
    crcCpm = 0
    strNum_Cpm = Space(6)
    i = ecf_NumeroCupom(strNum_Cpm)
    strCPF = strCNPJ
    If i = 0 Then
        MsgBox "Não foi possível encontrar o numero do cupom!", vbExclamation
        F_EmiteCupomVenda = False
        P_Status ""
        Exit Function
    End If
    If F_BuscaUmCampo("nota_sai", "tab_saidas", "num_ven='" & _
        sNum_ven & "' and serie_sai='CUPM' and cod_cai='" & strCod_Cai & "'") <> "" Then
        MsgBox "Este cupom já foi impresso!", vbExclamation
        Exit Function
    End If
    strSQL = "select   tv.*,ti.unit_ite,ti.desconto_ite,ti.qtd_ite,ti.seq_ite,(ti.unit_ite*ti.qtd_ite) as total, "
    strSQL = strSQL & "tp.* from   tab_produtos as tp, tab_itvendas as ti, tab_vendas as tv "
    strSQL = strSQL & " where  ti.cod_pro=tp.cod_pro and ti.num_ven='" & sNum_ven & "' "
    strSQL = strSQL & " and tv.num_ven=ti.num_ven "
    
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount = 0 Then
        MsgBox "Não existem itens nesta venda!", vbExclamation
        Exit Function
    End If
    sTipoMsg = F_BuscaUmCampo("tipomsgecf_par", "tab_Parametros", "")
    
    strInf = strInf & F_Centra(strNome_Cli, _
    48, "-") & Chr(13) & Chr(10)
    If Len(strCNPJ) = 18 Then
        strCPF = strCNPJ
    ElseIf Len(strCNPJ) = 14 Then
        strCPF = strCNPJ
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
        cValorPago = rPesq!valor_ven
    Else
        cValorPago = crcValorPago
    End If
    
    If bVendaEsp = True Then
        With rPesq
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
    
    

    With rPesq
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
                If !icms_pro = "01" Then
                    strST = "1700"
                ElseIf !icms_pro = "FF" Then
                    strST = "FF"
                Else
                    strST = Format(!icms_pro, "0#") & "00"
                End If
                
            ElseIf !st_pro = "R" Then
                strST = "FF" 'chamada para produto retido. (subst trib.)
            ElseIf !st_pro = "I" Then
                strST = "II" 'chamada para produto isento
            End If
            
            strUnidade = !unidade_pro
            
            crcQuantidade = !qtd_ite
            
            
            If bVendaEsp = False Then
                crcUnitario = !unit_ite
            Else
                crcUnitario = Round(!cf_pro + (!cf_pro * 0.3), 2)
            End If
            
            'crcDesconto = Round(!total / !Valor_ven * !desconto_ven, 2)
            crcDesconto = !desconto_ite
            'blnResp = F_VendaItem(strCod, strDescr, strST, _
            crcQuantidade, crcUnitario, strUnidade, crcDesconto)
            i = ecf_VendeItem(strCod, strDescr, strST, sTipoQtd, _
            IIf(sTipoQtd = "F", Format(crcQuantidade, "###,##0.000"), CStr(crcQuantidade)), 3, Format(crcUnitario, "###,##0.000"), "$", Format(crcDesconto, "###,##0.00"))
            If i = -2 Then
              
            ElseIf i <> 1 Then
                If i = -3 Then
                    MsgBox "O produto está cadastrado com o ICMS " & strST
                End If
                MsgBox "Erro na impressão do item!", vbExclamation
                   
            End If
            .MoveNext
        Loop
    End With
    
    If bTef Then
        F_EmiteCupomVenda = True
        Exit Function
    End If
    crcDesconto = 0 'arrumar o desconto
    
    i = ecf_IniciaFechamentoCupom("D", "$", CStr(crcDesconto))
    If i <> 1 Then
          MsgBox "Erro em ecf_iniciaFechamentoCupom, mod: = BancoDados Função F_EmiteCupomVenda", vbCritical
    End If
      
      
    i = ecf_EfetuaFormaPagamento("Dinheiro", Format(cValorPago, "###,##0.00"))
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

Sub P_GravaSaida(num_ven As String, nota_sai As String, Optional bVendaEsp As Boolean = False)
    Dim strSQL As String
    Dim sCod As String
    Dim sUnit As String
    Dim sDesc As String
    Dim sQtd As String
    Dim iCount As Integer
    Dim rPesq As New ADODB.Recordset
    Dim cValor As Currency
    On Error GoTo trata_Erro
    strSQL = "select * from tab_vendas where num_ven='" & num_ven & "' "
    abreObjetoAdo rPesq, strSQL
    nota_sai = "      "
    iCount = ecf_NumeroCupom(nota_sai)
    iCount = 0
    With rPesq
        strSQL = "Insert Into TAB_saidas ("
        strSQL = strSQL & "COD_CLI, NOTA_SAI, SERIE_SAI, EMISSAO_SAI, "
        strSQL = strSQL & "VALOR_SAI,"
        strSQL = strSQL & "DESCONTO_SAI, "
        strSQL = strSQL & "NUM_VEN,COD_USU, "
        strSQL = strSQL & "cod_cai,cod_tur,hora_sai) Values ( "
    
        strSQL = strSQL & "'" & !cod_cli & "', "
        strSQL = strSQL & "'" & F_SoNumero(Left(nota_sai, 6)) & "', "
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
    
        
    
        Conn.Execute strSQL
        .Close
    End With
    strSQL = "select * from tab_itvendas where num_ven='" & num_ven & "' order by seq_ite "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            iCount = iCount + 1 'sequencial
            strSQL = "insert into tab_itsai(cod_pro,nota_sai,serie_sai,unit_its,"
            strSQL = strSQL & "qtd_its,seq_its,cod_cai,icms_its)values("
            strSQL = strSQL & "'" & !cod_pro & "',"
            strSQL = strSQL & "'" & F_SoNumero(Left(nota_sai, 6)) & "',"
            strSQL = strSQL & "'CUPM',"
            
                strSQL = strSQL & F_GravaDec(!unit_ite) & ","
            
            strSQL = strSQL & F_GravaDec(!qtd_ite) & ","
            strSQL = strSQL & iCount & ","
            strSQL = strSQL & "'" & strCod_Cai & "', "
            strSQL = strSQL & F_GravaDec("07") & ")"
            
            
            Conn.Execute strSQL
            
            strSQL = "update tab_produtos set ef_pro=ef_pro-" & F_GravaDec(!qtd_ite)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            
            Conn.Execute strSQL
            'baixar também a saída
            .MoveNext
        Loop
        .Close
    End With
Exit Sub
trata_Erro:
    On Error Resume Next
    Open App.path & "\erro.txt" For Output As #2
        Print #2, Err.Description & Chr(13) + Chr(10)
        Print #2, strSQL
    Close #2
    MsgBox "Ocorreu um erro na gravação, e foi gerado o arquivo \erro.txt", vbExclamation
    Exit Sub
End Sub

Function F_AjustaSaldoCaixa(cod_cai As String, sData) As Currency
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cSaldo As Currency
    strSQL = "Select max(seq_cai) as max  from tab_caixa where data_cai <=" & F_GravaData(sData)
    strSQL = strSQL & "  and cod_cai='" & cod_cai & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            strSQL = "Select min(seq_cai) as max  from tab_caixa "
            strSQL = strSQL & "  where cod_cai='" & cod_cai & "' "
            abreObjetoAdo rPesq, strSQL
            strSQL = F_BuscaUmCampo("valor_cai", "tab_caixa", "seq_cai='" & _
                rPesq!max & "'")
        Else
            strSQL = F_BuscaUmCampo("saldo_cai", "tab_caixa", "seq_cai='" & _
                rPesq!max & "'")
            If strSQL = "" Then
                strSQL = "0"
            End If
            cSaldo = CCur(strSQL)
        
        
        End If
            
    End With
    strSQL = "select * from tab_caixa  where data_cai>=" & F_GravaData(sData)
    strSQL = strSQL & " and cod_Cai='" & cod_cai & "' order by data_cai,seq_cai "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If !tipo_cai = "C" Then
                cSaldo = cSaldo + !valor_Cai
            Else
                cSaldo = cSaldo - !valor_Cai
            End If
            strSQL = "update tab_caixa set saldo_cai=" & F_GravaDec(cSaldo)
            strSQL = strSQL & " where seq_Cai='" & !seq_cai & "'"
            Conn.Execute strSQL
            .MoveNext
        Loop
    End With
End Function


Function F_CalculaSaldo(data_fin As String, cod_con As String) As Currency
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cSaldo As Currency
    Dim bSaldo As Boolean
    On Error Resume Next
    strSQL = "select Max(data_fin) as max from tab_financeiro where data_fin < "
    strSQL = strSQL & F_GravaData(data_fin) & " and cod_Con='" & cod_con & "' "
    abreObjetoAdo rPesq, strSQL
    If IsNull(rPesq!max) Then
        strSQL = "select Min(data_fin) as max from tab_financeiro  where cod_Con='" & cod_con & "' "
        abreObjetoAdo rPesq, strSQL
        strSQL = "select min(seq_fin) as max from tab_financeiro where data_fin = "
        strSQL = strSQL & F_GravaData(rPesq!max) & " and cod_Con='" & cod_con & "' "
        abreObjetoAdo rPesq, strSQL
        bSaldo = False
    Else
        bSaldo = True
        strSQL = "select max(seq_fin) as max from tab_financeiro where data_fin = "
        strSQL = strSQL & F_GravaData(rPesq!max) & " and cod_Con='" & cod_con & "' "
        abreObjetoAdo rPesq, strSQL
'        strSQL = "select max(Seq_fin) as max from tab_financeiro where data_fin = "
'        strSQL = strSQL & F_GravaData(rPesq!Max)
'        rpesq.open strsql,conn
    End If
        strSQL = "select * from tab_financeiro where Seq_fin = '" & rPesq!max & "'" & " and cod_Con='" & cod_con & "' "
        abreObjetoAdo rPesq, strSQL
        cSaldo = rPesq!saldo_fin
        strSQL = "select * from tab_financeiro where data_fin >= " & F_GravaData(data_fin)
        strSQL = strSQL & " and cod_Con='" & cod_con & "' "
        strSQL = strSQL & " Order by  data_fin, Seq_fin"
        abreObjetoAdo rPesq, strSQL
        
        With rPesq
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
                Conn.Execute strSQL
            .MoveNext
            Loop
        End With
      
    F_CalculaSaldo = cSaldo
    
    
End Function

Sub P_MediaDiaria(sDataIni As String, sDataFin As String, Optional sMar As String, _
    Optional sCat As String, Optional sPro As String, Optional sFor As String = "")
                    
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim iDias As Integer
    Dim rPesq As New ADODB.Recordset
    Dim sMarca As String
    Dim sNome As String
    Dim sDias_Trab As String
    P_Status "Buscando produtos movimentados no período..."
    strSQL = "update tab_produtos set mmd_pro = 0"
    Conn.Execute strSQL
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

    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_produtos set mmd_pro = " & str(Round(!qtd / sDias_Trab, 4)) & ","
            strSQL = strSQL & "ipm_pro=" & F_GravaData(sDataIni) & ","
            strSQL = strSQL & "fpm_pro=" & F_GravaData(sDataFin)
            strSQL = strSQL & " where cod_pro = '" & !cod_pro & "' "
            
            Conn.Execute strSQL

            If .AbsolutePosition Mod 50 = 0 Or .AbsolutePosition = .RecordCount - 1 Then
                With frmMediaConsumo_Mer
                    .Image1.width = .Label3.width * ((rPesq.AbsolutePosition + 1) / rPesq.RecordCount)
                    .Image1.Refresh
                    .Label3.caption = Round((rPesq.AbsolutePosition + 1) * 100 / rPesq.RecordCount) & " %"
                End With
            End If
            .MoveNext
        Loop
        frmMediaConsumo_Mer.Label3.caption = .RecordCount & " Registros"
        .Close
    P_Status
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
    P_Status ""
    
End Sub
Function F_CalculaJuros(vencimento As String, data As String, valor As Currency, _
    Optional bJuroMaiorQueZero As Boolean = True) As Currency
    Dim sTxJuros As String
    Dim iDias As Integer
    On Error GoTo trata_Erro
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
trata_Erro:
    F_CalculaJuros = 0
End Function


Sub P_PreenchePlanoContas(obj As Object, optnCred As Boolean)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_planocontas  "
    If optnCred = True Then
        strSQL = strSQL & " where tipo_pla='C' "
    Else
        strSQL = strSQL & " where tipo_pla='D' "
    End If
    strSQL = strSQL & " order by nome_pla"
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    With rPesq
        Do While Not .EOF
            obj.AddItem rPesq!nome_pla
            .MoveNext
        Loop
    End With
    Exit Sub
trata_Erro:
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
        qtd_dev As Currency, data_dev As Date, gerado_Dev As Currency, tipo_dev As String, seq_pro As String) As Boolean

    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    'On Error GoTo trata_erro
    
    strSQL = "insert into tab_devolucoes(num_ven,cod_pro,qtd_ite,qtd_dev,data_dev, gerado_dev,"
    strSQL = strSQL & "cod_usu,seq_pro)Values("
    
    strSQL = strSQL & "'" & num_ven & "', "
    strSQL = strSQL & "'" & F_SoNumero(cod_pro) & "', "
    strSQL = strSQL & F_GravaDec(qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(qtd_dev) & ","
    
    strSQL = strSQL & F_GravaDataTempo(Now) & ",   "
    
    strSQL = strSQL & F_GravaDec(gerado_Dev) & ",   "
    strSQL = strSQL & "'" & strCod_Usu & "',  "
    strSQL = strSQL & seq_pro & ")"
    Conn.Execute strSQL
    
    strSQL = "update tab_produtos set er_pro=er_pro+"
    strSQL = strSQL & F_GravaDec(qtd_dev)
    strSQL = strSQL & " where cod_pro='" & cod_pro & "' "
    
    Conn.Execute strSQL
    
    
    
    
    F_GravaDevolucao = True
    
    Exit Function
trata_Erro:
    F_GravaDevolucao = False
End Function


Sub P_imprDevolucao(num_ven As String, obj As Object, Optional bVal As Boolean = True)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    
    'On Error GoTo trata_erro
    strSQL = "select tv.*,sum(tf.valor_ven) as valor, td.qtd_ite,td.tipo_dev,td.qtd_dev,td.data_dev,td.gerado_dev,"
    strSQL = strSQL & "tp.cod_pro , "
    strSQL = strSQL & "tp.nome_pro from tab_produtos as tp, tab_devolucoes as td,"
    strSQL = strSQL & "tab_vendas as tv,tab_venfec as tf "
    strSQL = strSQL & " where tv.num_ven=tf.num_ven and tv.num_ven=td.num_ven and "
    strSQL = strSQL & " td.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and tv.num_ven = '" & num_ven & "'"
    strSQL = strSQL & " and left(td.data_dev,10) = " & F_GravaData(Date) & ""
    strSQL = strSQL & " group by tf.num_ven order by nome_Pro "
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount = 0 Then
        MsgBox "Não há itens devolvidos nesta venda!", vbExclamation
        rPesq.Close
        Exit Sub
    End If
    With rPesq
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
            P_ImprX obj, "Valor: " & FormatVal(!valor), 160, False
        End If
        obj.fontsize = 10
        P_ImprLinha obj, "-"
        P_ImprX obj, "Código", 5
        P_ImprX obj, "Produto", 25
        P_ImprX obj, "T", 135
        P_ImprX obj, "Vendidos", 155, True, True
        P_ImprX obj, "Devolvidos", 180, True, True
        If bVal = True Then
            P_ImprX obj, "Vlr Gerado", 200, False, 3
        Else
            P_ImprX obj, "------", 200, False, 3
        End If
        P_ImprLinha obj, "-"
        Do While Not .EOF
            P_ImprX obj, Format(!cod_pro, "00#\.###\-#"), 5
            P_ImprX obj, !nome_pro, 25
            P_ImprX obj, !tipo_dev, 135
            P_ImprX obj, FormatVal(!qtd_ite), 155, True, True
            P_ImprX obj, FormatVal(!qtd_dev), 180, True, True
            If bVal Then
                P_ImprX obj, FormatVal(!gerado_Dev), 200, False, 3
            Else
                P_ImprX obj, "________", 200, False, 3
            End If
            cTotal = cTotal + !gerado_Dev
            .MoveNext
        Loop
        P_ImprLinha obj
        obj.fontsize = 14
        If bVal Then
            P_ImprX obj, "Total da Devolução............: " & FormatVal(cTotal), 200, False, 3
        Else
            P_ImprX obj, "Via do Entregador", 200, False, 3
        End If
    End With
    
    obj.EndDoc
    Exit Sub
trata_Erro:
    Printer.KillDoc
End Sub


Sub P_PreencheBarCode(cod_pro As String, obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select * from tab_codbarras where cod_pro='" & cod_pro & "' order by cod_bar"
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    DoEvents
    With rPesq
        Do While Not .EOF
            obj.AddItem !cod_bar
            .MoveNext
        Loop
        
        .Close
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_PreencheEmbalagens(cod_pro As String, obj As Object, Optional bUN As Boolean = False)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select * from tab_embalagens where cod_pro='" & cod_pro & "' order by qtd_emb"
    abreObjetoAdo rPesq, strSQL
    obj.Clear
    DoEvents
    
    With rPesq
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
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Function F_IcmsProd(cod_pro As String, Retorno As Integer) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro:
    strSQL = "select icms_pro from tab_produtos where cod_pro='" & cod_pro & "'"
    
    
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            If Retorno = 3 Then
                F_IcmsProd = !perc_imp
            ElseIf Retorno = 2 Then
                F_IcmsProd = !desc_imp
            ElseIf Retorno = 1 Then
                F_IcmsProd = !ecf_imp
            Else
                F_IcmsProd = !cod_imp
            End If
        Else
            MsgBox "O imposto do produto não está cadastrado!", vbExclamation
            frmCadPro.tag = cod_pro
            frmCadPro.Show 1
        End If
        .Close
    End With
    Exit Function
trata_Erro:
    
 End Function
Function F_TaxasEntradas() As Currency
    F_TaxasEntradas = crcTxFederativa_Par + crcTxAdm + crcTxMargemSeg_Par
End Function
'Function F_CancelaVendaAposCupom(sNum_ven As String) As Boolean
'    Dim iPos As Integer
'    Dim strSQL As String
'    Dim rPesq as new adodb.recordset
'    On Error GoTo trata_Erro
'    'varre toda a venda voltando o estoque real dos itens depois muda para situacao 5
'    F_CancelaVendaAposCupom = False
'
'
'    If F_BuscaUmCampo("status_ven", "tab_vendas", "num_ven='" & sNum_ven & "'") = "3" Then
'        MsgBox "Não foi possível cancelar os lançamentos da venda, pois ela já se encontra cancelada!", vbExclamation
'        F_CancelaVendaAposCupom = False
'        Exit Function
'    End If
'
'
'
'    strSQL = "select * from tab_itvendas where num_ven='" & sNum_ven & "'  "
'    rpesq.open strsql,conn
'    With rPesq
'        Do While Not .EOF
'
'
'
'
'
'
'
'            strSQL = "update tab_PRODUTOS set er_pro=er_pro + " & F_GravaDec(CStr(!qtd_ite))
'            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
'
'            conn.execute strSQL
'
'            .MoveNext
'        Loop
'        .Close
'    End With
'    strSQL = "select * from tab_venfec where num_ven='" & sNum_ven & "' "
'    rpesq.open strsql,conn
'    With rPesq
'        Do While Not .EOF
'
'            strSQL = F_GravaCai("D", Date, CStr(!valor_ven), "Venda " & !num_ven & " Cancelada", !num_ven, _
'                F_BuscaUmCampo("tipo_cai", "tab_formaspag", "cod_for='" & !cod_for & "'"), "01", _
'                0, 0, strCod_Cai, strCod_Tur, "000")
'
'
'            .MoveNext
'        Loop
'        .Close
'    End With
'    strSQL = "select * from tab_doc where num_ven='" & sNum_ven & "' and esp_doc='V' "
'                rpesq.open strsql,conn
'                With rPesq
'                    Do While Not .EOF
'                        strSQL = "delete from tab_itdoc where seq_doc='" & !seq_doc & "'"
'                        conn.execute strSQL
'                        strSQL = "delete from tab_cheques where seq_doc='" & !seq_doc & "'"
'                        conn.execute strSQL
'                        strSQL = "delete from tab_doc where seq_doc='" & !seq_doc & "'"
'                        conn.execute strSQL
'                        .MoveNext
'                    Loop
'                End With
'
'    strSQL = "update tab_vendas set situacao_ven=5, "
'    strSQL = strSQL & " cancelavenda_usu='" & strCod_Usu & "' "
'    strSQL = strSQL & " where num_ven='" & sNum_ven & "' "
'    conn.execute strSQL '5 = venda cancelada
'    F_CancelaVendaAposCupom = True
'    Exit Function
'trata_Erro:
'    P_MostraErro
'End Function


Function F_RetornaLimiteUsoCli(cod_cli As String) As Currency
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select sum(ti.reman_itd) as total "
    strSQL = strSQL & " from tab_itdoc as ti, tab_doc as td"
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.cod_ter='" & cod_cli & "' "
    strSQL = strSQL & " and td.tipo_doc='R' "
    abreObjetoAdo rPesq, strSQL
    If Not IsNull(rPesq!total) Then
        F_RetornaLimiteUsoCli = rPesq!total
    End If
    
    
    strSQL = "select sum(ti.valor_che) as total "
    strSQL = strSQL & " from tab_cheques as ti, tab_doc as td"
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.cod_ter='" & cod_cli & "' "
    strSQL = strSQL & " and td.tipo_doc='R' "
    strSQL = strSQL & " and ti.dep_che is null "
    abreObjetoAdo rPesq, strSQL
    If Not IsNull(rPesq!total) Then
        F_RetornaLimiteUsoCli = rPesq!total + F_RetornaLimiteUsoCli
    End If
    
    
    
    rPesq.Close
End Function


Sub P_CriaTabTmpTER()
    Dim strSQL As String
    strSQL = "drop table if exists tmp_01;"
    Conn.Execute strSQL
    strSQL = strSQL & "create table tmp_01(cod_ter char(5) not null,"
    strSQL = strSQL & "nome_ter char(5) not null,"
    strSQL = strSQL & "tipo_ter char(1) not null,"
    strSQL = strSQL & "primary key (cod_ter,tipo_ter));"
    Conn.Execute strSQL
    strSQL = strSQL & "insert into tmp_01 select cod_cli,nome_cli,'C' as tipo from tab_clientes;"
    Conn.Execute strSQL
    strSQL = strSQL & "insert into tmp_01 select cod_rep,nome_rep,'R' as tipo from tab_representantes;"
    Conn.Execute strSQL
End Sub
Function VerificaDataBanco() As Boolean
    Dim strSQL As String
    Dim dataBanco As String
    Dim rPesq As New ADODB.Recordset
    Dim difDi As Integer
    Dim difHo As Integer
    VerificaDataBanco = True
    strSQL = "select now() as data"
    abreObjetoAdo rPesq, strSQL
    
    'strSQL = F_BuscaUmCampo("now()")
    strSQL = rPesq!data
 
    difDi = DateDiff("D", CDate(strSQL), Now)
    difHo = DateDiff("H", CDate(strSQL), Now)
    If Int(difDi) <> 0 Then
        MsgBox "Há uma diferença de datas de " & difDi & _
        " dias entre a data do Gerenciador de banco de dados (" & strSQL & ") " & _
        " e a data deste dispositivo (" & Now & ")", vbExclamation
        VerificaDataBanco = False
    ElseIf (difHo > 2 Or difHo < -2) And difHo <> 0 Then
        MsgBox "Há uma diferença de datas de " & difHo & _
        " HORAS entre a data do Gerenciador de banco de dados (" & strSQL & ") " & _
        " e a data deste dispositivo (" & Now & ")", vbExclamation
        VerificaDataBanco = False
    End If
    
    
End Function
Function F_AbreBancoODBCAdo(c As ADODB.Connection, nome As String) As Boolean
    
    On Error GoTo erro
    'DBEngine.DefaultType = dbUseODBC
    Set c = New ADODB.Connection
    
    c.Open nome
    c.CursorLocation = adUseClient
    
    
   
    F_AbreBancoODBCAdo = True
    Exit Function
erro:
    MsgBox "Erro na Abertura do Banco!", vbExclamation
    F_AbreBancoODBCAdo = False
End Function

Sub P_PreencheSituacaoDoc(obj As Object)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from sped_situacao_doc order by codigo "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            obj.AddItem rPesq!codigo & "." & rPesq!descricao
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub
Sub P_PreencheModeloNFTab411(obj As Object)
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from sped_tipo_doc_icms order by codigo "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            obj.AddItem rPesq!codigo & "." & rPesq!descricao
            .MoveNext
        Loop
        .Close
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub

Sub P_cancelanotaFechada(seq_not As String, Optional tipo As String = "F")
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTab As String
    Dim sTabIte As String
    Dim sEst As String
    
    sTab = "sped_nf"
    sTabIte = "sped_nf_itens"
    sEst = "er_pro"
    
    
    strSQL = "select tc.*, ti.* from sped_nf as tc inner join sped_nf_itens as ti on tc.id=ti.id_nf"
    strSQL = strSQL & " where tc.id='" & seq_not & "'"
    
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update " & TAB_PRODUTOS & " set er_pro=er_pro-" & F_GravaDec(!qtd_convert)
            If !tipo_entrada = "01" Then
                strSQL = strSQL & ",ef_pro=ef_pro-" & F_GravaDec(!qtd_convert)
            End If
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Conn.Execute strSQL
            
            
            If bUsaEstoqueSeparado Then
                strSQL = "update estoque set er=er-" & F_GravaDec(!qtd_convert)
                If !tipo_entrada = "01" Then
                    strSQL = strSQL & ",ef=ef-" & F_GravaDec(!qtd_convert)
                End If
                strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
                strSQL = strSQL & " and cod_emp='" & !cod_emp & "'"
                Conn.Execute strSQL
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "update " & sTab & " set status=5 where id='" & seq_not & "' "
    
    Conn.Execute strSQL
    
    MsgBox "Nota cancelada, Estoque reposto!", vbExclamation
    
End Sub


Function Encrypt(senha) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select aes_encrypt('" & senha & "','" & encryptKey & "') as senha"
    abreObjetoAdo rPesq, strSQL
    Encrypt = rPesq!senha
    rPesq.Close
    
End Function
Function Decrypt(senha) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select aes_encrypt('" & senha & "','" & encryptKey & "') as senha"
    abreObjetoAdo rPesq, strSQL
    Decrypt = rPesq!senha
    rPesq.Close
    
End Function
