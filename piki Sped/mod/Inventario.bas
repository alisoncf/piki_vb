Attribute VB_Name = "inventario"
Dim CrLf As String
Type inventario
    id As String
    data As String
    cod_emp As String
End Type
Function SelecionaInventario() As inventario.inventario
    Dim fr As New frmBuscaInv
    
    fr.tag = "x"
    fr.Show 1
    If fr.tag <> "" Then
        SelecionaInventario.id = fr.ado!id
        SelecionaInventario.data = fr.ado!data
        SelecionaInventario.cod_emp = fr.ado!cod_emp
    End If
    Unload fr
    
End Function
Sub AlteraInventario(id As String, de As String, ate As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "update inventario set "
    strSQL = strSQL & "de=" & F_GravaData(de) & ","
    strSQL = strSQL & "ate=" & F_GravaData(ate) & " "
    strSQL = strSQL & " where id=" & id
    conn.Execute strSQL
Exit Sub
trata_Erro:
    P_MostraErro
End Sub
    

Function geraInventario(data As String, de As String, ate As String, cod_emp As String) As String
    Dim strSQL As String
    Dim id As String
    On Error GoTo trata_Erro
    id = MaxKeySQL("inventario", "id", "", "")
    strSQL = "insert into inventario(id,data,de,ate,cod_emp)values("
    strSQL = strSQL & id & ","
    strSQL = strSQL & F_GravaDataTempo(data) & ","
    strSQL = strSQL & F_GravaData(de) & ","
    strSQL = strSQL & F_GravaData(ate) & ","
    strSQL = strSQL & "'" & cod_emp & "') "
    conn.Execute strSQL
    geraInventario = id
trata_Erro:
    
End Function

Sub grava_item(id_invent As String, codigo As String, nome As String, _
    custo As String, preco As String, q1 As String, q2 As String, cod_emp As String)
    
    Dim strSQL As String
    strSQL = "insert into inventario_it(inventario,codigo,nome,custo,preco,q1,q2,cod_emp)values("
    strSQL = strSQL & id_invent & ","
    strSQL = strSQL & "'" & codigo & "',"
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & F_GravaDec(custo) & ","
    strSQL = strSQL & F_GravaDec(preco) & ","
    strSQL = strSQL & F_GravaDec(q1) & ","
    strSQL = strSQL & F_GravaDec(q2) & ","
    strSQL = strSQL & "'" & cod_emp & "') "
    conn.Execute strSQL
    
End Sub

Sub GeraInventarioSintegra(inventario As String, dataI As String, dataF As String, Optional caminho As String = "")

    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim buf As Variant
    Dim s75 As String
    Dim total74 As Currency
    Dim total75 As Currency
    Dim Itens74 As Integer
    Dim Itens75 As Integer
    Dim arq1 As String
    Dim arq2 As String
    Dim espacos As String
    On Error GoTo trata_Erro
    CrLf = Chr(13) + Chr(10)
    arq74 = App.Path & "\tmp74.txt"
    arq75 = App.Path & "\tmp75.txt"
    
    
    
    If caminho = "" Then
        caminho = App.Path & "\" & "inv_" & GravaDataSintegra(dataI) & "_" & GravaDataSintegra(dataF) & Left(strNome_Emp, 10) & ".txt"
    End If
    
    If Dir(caminho) <> "" Then
        If MsgBox("O arquivo " & caminho & " já existe no caminho especificado. Deseja substituí-lo?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        
        End If
    End If
    
    
    If Not AbreArquivoExterno(caminho, 1) Then
        MsgBox "Não foi possível abrir o caminho " & caminho, vbExclamation
        Exit Sub
    End If
    
    
    
    strSQL = "select ti.*,tp.nome_pro,tp.ncm_pro,tp.unidade_pro,tp.icms_pro"
    strSQL = strSQL & " from inventario_it as ti, tab_produtos as tp "
    strSQL = strSQL & " where tp.cod_pro=ti.codigo "
    strSQL = strSQL & " and ti.inventario=" & inventario
    strSQL = strSQL & " and q1>0 and custo>0 "
    strSQL = strSQL & " order by tp.cod_pro "
    
    abreObjetoAdo rPesq, strSQL
    buf = GeraTipo10(strCNPJ_Emp, strIE_Emp, strNome_Emp, strCidade_Emp, strUF_Emp, strFax_Emp, dataI, dataF, "3", "1", "1") & CrLf
    buf = geraTipo10Padrao(dataI, dataF)
    ImprimeEmArquivoExterno buf, 1
    
    buf = geraTipo11Padrao
    ImprimeEmArquivoExterno buf, 1
    With rPesq
        Itens74 = .RecordCount
        Itens75 = .RecordCount
        Do While Not .EOF
            buf = GeraItemTipo74(dataF, !codigo, !q1, CStr(!custo * !q1), DoInformanteEmSeuPoder, "00000000000000", "", _
            strUF_Emp)
            
            
            
            
            ImprimeEmArquivoExterno buf, 1
            
            s75 = s75 & GeraItemTipo75(dataI, dataF, !codigo, Format(!ncm_pro, "0000000#"), !nome_pro, !unidade_pro, "0", !icms_pro, _
            "0", "0") & CrLf
            total74 = total74 + (!custo * !q1)
            .MoveNext
        Loop
        
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            
            buf = GeraItemTipo75(dataI, dataF, !codigo, Format(!ncm_pro, "0000000#"), !nome_pro, !unidade_pro, "0", !icms_pro, _
            "0", "0")
            ImprimeEmArquivoExterno buf, 1
            
            
            
            .MoveNext
        Loop
        
        
        .Close
    End With
    
    
    
    buf = GeraItemTipo90_74(strCNPJ_Emp, strIE_Emp, CStr(Itens74), CStr(Itens75)) & GeraItemTipo99(Itens74 + Itens75 + 3)
    
    If Len(buf) < 125 Then
        espacos = EsticaTexto(" ", 125 - Len(buf))
    End If
    buf = buf & espacos & "1"
    
    
        ImprimeEmArquivoExterno buf, 1



    'GravaArquivoExterno caminho, buf
    If FechaArquivoExterno(1) Then
        If MsgBox("Arquivo gerado em " & caminho & "." & CrLf & "Deseja abrir arquivo?", vbYesNo + vbQuestion) = vbYes Then
            Shell "explorer " & caminho
        End If
    Else
    
    End If

    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub RetornaInventario(inventario As String, data As String)
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select ti.cod_pro,it.q1,it.q2,sum(ti.qtd_ite) as total from tab_vendas as tv, tab_itvendas as ti, inventario_it as it "
    strSQL = strSQL & " where tv.num_ven=ti.num_ven "
    strSQL = strSQL & " and ti.cod_pro=it.codigo "
    strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(data)
    strSQL = strSQL & " and tv.situacao_ven=4 "
    strSQL = strSQL & " and ti.inventario=" & inventario
    strSQL = strSQL & " group by ti.cod_pro "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            .MoveNext
        Loop
    End With
    
    
    
End Sub
