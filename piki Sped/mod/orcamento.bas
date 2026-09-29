Attribute VB_Name = "orcamento"
Enum status_orcamento
    emaberto = 1
    fechado = 2
    cancelado = 9
End Enum
Type orcamento

    id As String
    cliente As String
    cod_cli As String
    valor As String
    telefone As String
    contato As String
    email As String

End Type
Type detalhamento_uniforme
    cor As String
    silk As String
    bordado As String
    detalhamento As String
    Flag As Boolean
End Type


Function exclui_orcamento_item(id As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = " delete from orcamento_itens where id=" & id
    conn.Execute strSQL
    exclui_orcamento_item = True
trata_Erro:
    
End Function

Function AlteraOrcamentoStatus(id As String, status As status_orcamento) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = " update orcamentos set status = " & status & "  where id=" & id
    conn.Execute strSQL
    AlteraOrcamentoStatus = True
trata_Erro:
    
End Function

Function insereOrcamento_item(id_orcamento As String, cod_pro As String, Produto As String, qtd As String, unit As String, _
cor As String, silk As String, bordado As String, detalhamento As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    On Error GoTo trata_Erro
    
    strSQL = "insert into orcamento_itens(id_orcamento,cod_pro,produto,quantidade,unitario,total,cor,silk,bordado,detalhamento)values("
    strSQL = strSQL & "'" & id_orcamento & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & "'" & Produto & "',"
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & F_GravaDec(unit) & ","
    strSQL = strSQL & F_GravaDec(qtd * unit) & ","

    strSQL = strSQL & "'" & cor & "',"
    strSQL = strSQL & "'" & silk & "',"
    strSQL = strSQL & "'" & bordado & "',"
    strSQL = strSQL & "'" & detalhamento & "')"
    conn.Execute strSQL
    
    
    strSQL = "select max(id) as id from orcamento_itens"
    abreObjetoAdo rPesq, strSQL
    
    
    insereOrcamento_item = rPesq!id
    rPesq.Close
    
    
Exit Function
trata_Erro:
    
End Function

Function AlteraOrcamento_item(id As String, cod_pro As String, qtd As String, unit As String, _
cor As String, silk As String, bordado As String, detalhamento As String) As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    On Error GoTo trata_Erro
    strSQL = "replace into orcamento_itens(id,cod_pro,quantidade,unitario,total,descricao ],cor,silk,bordado,detalhamento)values("
    strSQL = strSQL & "'" & id_orcamento & "',"
    strSQL = strSQL & F_GravaDec(qtd) & "',"
    strSQL = strSQL & F_GravaDec(unit) & "',"
    strSQL = strSQL & F_GravaDec(qtd * unit) & "',"
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & "'" & cor & "',"
    strSQL = strSQL & "'" & silk & "',"
    strSQL = strSQL & "'" & bordado & "',"
    strSQL = strSQL & "'" & detalhamento & "' )"
    
    conn.Execute strSQL
    
    
    
    
    
    AlteraOrcamento_item = True


    
Exit Function
trata_Erro:
    
End Function


Function insereOrcamento(cod_cli As String, cliente As String, telefone As String, contato As String, email As String, cod_For As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If cod_cli = "" Then
        cod_cli = "0"
    End If
    If cod_For = "" Then
        cod_For = "01"
    End If

    On Error GoTo trata_Erro
  
    strSQL = "insert into orcamentos(cod_cli,telefone,cliente,contato,email,cod_for,data,cod_usu)values("
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & "'" & telefone & "',"
    strSQL = strSQL & "'" & cliente & "',"
    strSQL = strSQL & "'" & contato & "',"
    strSQL = strSQL & "'" & email & "',"
    strSQL = strSQL & "'" & cod_For & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "')"
    conn.Execute strSQL
    
    
    strSQL = "select max(id) from orcamentos limit 1"
    abreObjetoAdo rPesq, strSQL
    
    
    insereOrcamento = rPesq.Fields(0).Value
    rPesq.Close

    
Exit Function
trata_Erro:
    
End Function
Sub ExcluiOrcamento(id As String, Optional permanente As Boolean = False)
    Dim strSQL As String
    If F_BuscaUmCampo("status", "orcamento", "id=" & id) = "2" Then
        MsgBox "O orçamento não pode ser excluído uma vez que foi efetivado!", vbExclamation
        Exit Sub
    End If
    
    
    If permanente Then
        strSQL = "delete from orcamentos where id=" & id
        conn.Execute strSQL
        
        strSQL = "delete from orcamento_itens where id=" & id
        conn.Execute strSQL
    Else
    
        strSQL = "update orcamentos set status=9 where id=" & id
        conn.Execute strSQL
    End If
    
    
End Sub

Function duplicarOrcamento(id As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim id_novo As String
    strSQL = "select * from orcamentos where id=" & id
    abreObjetoAdo rPesq, strSQL
    With rPesq
        id_novo = insereOrcamento(!cod_cli, !cliente, !telefone, !contato, !email, !cod_For)
             
        
        .Close
        
    End With
    Dim item As String
    strSQL = "select * from orcamento_itens where id_orcamento=" & id
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            
            item = orcamento.insereOrcamento_item(id_novo, !cod_pro, TiraNull(!Produto), !Quantidade, !unitario, TiraNull(!cor), TiraNull(!silk), TiraNull(!bordado), TiraNull(!detalhamento))
            .MoveNext
        Loop
        .Close
    End With
    
    
    
    
    
    duplicarOrcamento = id_novo
    
    
End Function

Function AlteraOrcamento(id As String, cod_cli As String, cliente As String, telefone As String, _
    cod_For As String, valor_final As String, desconto As String, contato As String, email As String, Optional obs As String = "") As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim i As Integer
    If cod_cli = "" Then
        cod_cli = "0"
    End If
    If cod_For = "" Then
        cod_For = "01"
    End If

    On Error GoTo trata_Erro
    strSQL = "update orcamentos set "
    strSQL = strSQL & "cod_cli='" & cod_cli & "',"
    strSQL = strSQL & "telefone='" & telefone & "',"
    strSQL = strSQL & "cliente='" & cliente & "',"
    strSQL = strSQL & "email='" & email & "',"
    strSQL = strSQL & "observacao='" & obs & "',"
    strSQL = strSQL & "contato='" & contato & "',"
    strSQL = strSQL & "cod_for='" & cod_For & "', "
    strSQL = strSQL & "data=" & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "desconto=" & F_GravaDec(desconto) & ","
    
    strSQL = strSQL & "valor_final=" & F_GravaDec(valor_final) & " "
    strSQL = strSQL & " where id=" & id
    conn.Execute strSQL
    
    
    AlteraOrcamento = True
    
        
    


    
Exit Function
trata_Erro:
    
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
    
    conn.Execute strSQL
    F_GeraNovoOrc = sNum_ven
    
    
    
End Function
Sub ImprimeOrcamento(id As String)
    ImprimeOrcamento_Arquivo id
End Sub

Sub ImprimeOrcamento1(id As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    strSQL = "select o.id as numero_orcamento,fp.nome_for, o.*,c.nome_cli,c.cnpj_cli, i.*, f.guerra_fun from orcamentos "
    strSQL = strSQL & " as o inner join clientes as c on o.cod_cli=c.cod_cli "
    strSQL = strSQL & " left join " & TAB_FUNCIONARIOS & " as f on o.cod_usu=f.cod_fun "
    strSQL = strSQL & " inner join " & TAB_FORMASPAG & " as  fp on fp.cod_for=o.cod_for "
    strSQL = strSQL & " left join orcamento_itens as i on i.id_orcamento=o.id"
    strSQL = strSQL & " where 1>0 "
    strSQL = strSQL & " and o.id=" & id
    abreObjetoAdo rPesq, strSQL
    
    
    Dim s As String
    Dim path As String
    
    path = html.F_AbreHtml("Orcamento_" & id)
    html.P_HtmlAbreHead 1
    html.AbreTitle 1, "Orçamento " & id
    html.FechaHead
    html.AbreBody 1
    html.AbreTable 1, "style=" & entreAspas("width:100%;border: 0px solid")
    html.AbreLinha 1
    html.AbreColuna 1, "style=" & entreAspas("text-align:center")
    If Dir("../pic/logo_topo.jpg") <> "" Then
        html.add_figura 1, "../pic/logo_topo.jpg"
    Else
        html.add_figura 1, "../pic/logo_topo.png"
    End If
    html.FechaColuna 1
    html.FechaTable
    
    html.AbreDiv 1
        html.AbreTable 1, "style=" & entreAspas("width:100%;border: 0px solid; font-family:Arial ; font-size:28px")
            html.AbreLinha 1
                html.AddColuna 1, "ORÇAMENTO Nº " & rPesq!numero_orcamento, "style=" & entreAspas("text-align:center")
                Print #1, "<br>"
            html.FechaLinha 1
        html.FechaTable
    html.FechaDiv 1
    
    html.AbreDiv 1
        html.LinhaHorizontal
    html.FechaDiv 1
    
    
    html.AbreDiv 1
        html.AbreTable 1, "style=" & entreAspas("width:100%;border: 0px solid; font-family:Arial ; font-size:20px; font-weitght:bold")
            html.AbreLinha 1
                html.AddColuna 1, "Cliente: " & rPesq!cliente, "style=" & entreAspas("text-align:left")
                html.AddColuna 1, "Data: " & Format(rPesq!data, "dd/mm/yyyy") & " Validade: 30 dias", "style=" & entreAspas("text-align:left")
            html.FechaLinha
        html.FechaTable
    html.FechaDiv 1
    
    html.AbreDiv 1
        html.LinhaHorizontal
    html.FechaDiv 1
    
    html.AbreDiv 1
        html.AbreTable 1, "style=" & entreAspas("width:100%;border: 1px solid; font-family:Arial ; font-size:18px;border-collapse:collapse")
            html.AbreLinha 1
                html.AddColuna 1, "Qtde: ", "style=" & entreAspas("text-align:left;border: 1px solid; font-weight: bold")
                html.AddColuna 1, "Produto: ", "style=" & entreAspas("text-align:left;border: 1px solid; font-weight: bold")
                html.AddColuna 1, "Detalhamento: ", "style=" & entreAspas("text-align:left;border: 1px solid; font-weight: bold")
                html.AddColuna 1, "Cor: ", "style=" & entreAspas("text-align:left;border: 1px solid; font-weight: bold")
                html.AddColuna 1, "Silk: ", "style=" & entreAspas("text-align:left;border: 1px solid; font-weight: bold")
                html.AddColuna 1, "Bordado: ", "style=" & entreAspas("text-align:left;border: 1px solid; font-weight: bold")
                html.AddColuna 1, "Vlr Unitário: ", "style=" & entreAspas("text-align:left;border: 1px solid")
                html.AddColuna 1, "Total: ", "style=" & entreAspas("text-align:left;border: 1px solid")
            html.FechaLinha
        
            With rPesq
                Do While Not .EOF
                    html.AbreLinha 1
                        html.AddColuna 1, TiraNull(!Quantidade), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                        html.AddColuna 1, TiraNull(!Produto), "style=" & entreAspas("text-align:left;border: 1px solid")
                        html.AddColuna 1, TiraNull(!detalhamento), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                        html.AddColuna 1, TiraNull(!cor), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                        html.AddColuna 1, TiraNull(!silk), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                        html.AddColuna 1, TiraNull(!bordado), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                        html.AddColuna 1, FormatVal(!unitario), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                        html.AddColuna 1, FormatVal(!Quantidade * !unitario), "style=" & entreAspas("text-align:left;border: 1px solid;border-collapse:collapse")
                    html.FechaLinha
                    If Not IsNull(!unitario * !Quantidade) Then
                        cTotal = cTotal + (!unitario * !Quantidade)
                    End If
                    .MoveNext
                Loop
                .MoveFirst
            End With
        
        html.FechaTable
    html.FechaDiv 1
    
    
    html.AbreDiv 1
        html.AbreTable 1, "style=" & entreAspas("width:100%;border: 1px solid; font-family:Arial ; font-size:18px; background-color: #F2F2F2")
            html.AbreLinha 1
                html.AbreColuna 1, "style=" & entreAspas("text-align:center;  1px solid")
                html.AbreP 1, "Condições de Pagamento", "style=" & entreAspas("text-align:center")
                html.AbreP 1, rPesq!nome_For, "style=" & entreAspas("text-align:center")
                html.FechaColuna 1
                
                html.AbreColuna 1, "style=" & entreAspas("text-align:center")
                html.AbreP 1, "Total:", "style=" & entreAspas("text-align:center")
                html.AbreP 1, "R$" & FormatVal(cTotal), "style=" & entreAspas("text-align:center")
                html.FechaColuna 1
                
                html.AbreColuna 1, "style=" & entreAspas("text-align:center")
                html.AbreP 1, "Desconto:", "style=" & entreAspas("text-align:center")
                html.AbreP 1, "R$" & FormatVal(rPesq!desconto), "style=" & entreAspas("text-align:center")
                html.FechaColuna 1
                
                html.AbreColuna 1, "style=" & entreAspas("text-align:center")
                html.AbreP 1, "Valor Final:", "style=" & entreAspas("text-align:center; font-weight: bold")
                html.AbreP 1, "R$ " & FormatVal(rPesq!valor_final), "style=" & entreAspas("text-align:center; font-weight: bold")
                html.FechaColuna 1
                
                
                html.LinhaHorizontal 1
                
                html.LinhaHorizontal 1
            html.FechaLinha
        html.FechaTable
    html.FechaDiv 1
    
    
    
    html.AbreDiv 1
        html.AbreTable 1, "style=" & entreAspas("width:100%;border: 1px solid; font-family:Arial ; font-size:14px")
            html.AbreLinha 1
                html.AddColuna 1, strNome_Emp & " - " & strCidade_Emp & "/" & strUF_Emp & "CNPJ:" & strCNPJ_Emp & "  ", "style=" & entreAspas("text-align:left")
                html.AddColuna 1, Now, "style=" & entreAspas("text-align:right")
                html.LinhaHorizontal 1
                
            html.FechaLinha
        html.FechaTable
    html.FechaDiv 1
    
    html.FechaBody 1
    
    html.P_FechaHtml 1, True, path
    
    
    
    
End Sub
Sub ImprimeOrcamento_Arquivo(id As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    Dim nomeArquivo As String
    Dim s As String
    Dim path As String
    Dim conteudo As String
    Dim itens As String
    Dim caminho_modelo As String
    Dim caminho_orcamento As String
    strSQL = "select o.id as numero_orcamento,fp.nome_for, o.*,c.nome_cli,c.cnpj_cli, i.*, f.guerra_fun from orcamentos "
    strSQL = strSQL & " as o inner join clientes as c on o.cod_cli=c.cod_cli "
    strSQL = strSQL & " left join " & TAB_FUNCIONARIOS & " as f on o.cod_usu=f.cod_fun "
    strSQL = strSQL & " inner join " & TAB_FORMASPAG & " as  fp on fp.cod_for=o.cod_for "
    strSQL = strSQL & " left join orcamento_itens as i on i.id_orcamento=o.id"
    strSQL = strSQL & " where 1>0 "
    strSQL = strSQL & " and o.id=" & id
    abreObjetoAdo rPesq, strSQL
    
    
    caminho_modelo = App.path & "\pic"
    caminho_logo = caminho_modelo & "\logo_topo.png"
    caminho_orcamento = App.path & "\html\orcamentos"
    If Dir(caminho_orcamento, vbDirectory) = "" Then
         
        MkDir caminho_orcamento
    End If
    FileCopy caminho_logo, caminho_orcamento & "\logo_topo.png"
    
    conteudo = LeArquivoExternoUTF8(caminho_modelo & "\modelo_orcamento.html")
        
        conteudo = Replace(conteudo, "#titulo", "Orçamento Nº " & rPesq!id)
        
        conteudo = Replace(conteudo, "#data", rPesq!data)
        conteudo = Replace(conteudo, "#cliente", rPesq!nome_Cli)
        
        itens = ""
            With rPesq
                Do While Not .EOF
                    itens = itens & "<tr>"
                    itens = itens & "<td>" & TiraNull(!Quantidade) & "</td>"
                        itens = itens & "<td>" & TiraNull(!Produto) & "</td>"
                        itens = itens & "<td>" & TiraNull(!detalhamento) & "</td>"
                        itens = itens & "<td>" & TiraNull(!cor) & "</td>"
                        itens = itens & "<td>" & TiraNull(!silk) & "</td>"
                        itens = itens & "<td>" & TiraNull(!bordado) & "</td>"
                        itens = itens & "<td>" & FormatVal(!unitario) & "</td>"
                        itens = itens & "<td>" & FormatVal(!Quantidade * !unitario) & "</td>"
                        itens = itens & "</tr>"

                        
                    If Not IsNull(!unitario * !Quantidade) Then
                        cTotal = cTotal + (!unitario * !Quantidade)
                    End If
                    .MoveNext
                Loop
                .MoveFirst
            End With
        conteudo = Replace(conteudo, "#itens", itens)
        conteudo = Replace(conteudo, "#total", FormatVal(cTotal))
        conteudo = Replace(conteudo, "#desconto", FormatVal(rPesq!desconto))
        conteudo = Replace(conteudo, "#valor_final", FormatVal(rPesq!valor_final))
        conteudo = Replace(conteudo, "#empresa", strNome_Emp)
        conteudo = Replace(conteudo, "#cidade", strCidade_Emp)
        conteudo = Replace(conteudo, "#uf", strUF_Emp)
        conteudo = Replace(conteudo, "#cnpj", strCNPJ_Emp)
        conteudo = Replace(conteudo, "#datahora", Now)
    
        
        nomeArquivo = caminho_orcamento & "\orcamento" & rPesq!id & "_" & rPesq!nome_Cli & ".html"
        GravaArquivoExternoUTF8 nomeArquivo, conteudo, True
        rPesq.Close
    
End Sub

Function GeraVenda(id_orcamento As String, Optional cod_emp As String = "01") As String
    Dim strSQL As String
    Dim num_ven As String
    Dim rPesq As New ADODB.Recordset
    Dim rItens As New ADODB.Recordset
    On Error GoTo trata_Erro
    
    
    strSQL = "select o.id as numero_orcamento,fp.nome_for, o.*,c.nome_cli,c.cnpj_cli, f.guerra_fun from orcamentos "
    strSQL = strSQL & " as o inner join clientes as c on o.cod_cli=c.cod_cli "
    strSQL = strSQL & " left join " & TAB_FUNCIONARIOS & " as f on o.cod_usu=f.cod_fun "
    strSQL = strSQL & " inner join " & TAB_FORMASPAG & " as  fp on fp.cod_for=o.cod_for "
    
    strSQL = strSQL & " where 1>0 "
    strSQL = strSQL & " and o.id=" & id_orcamento
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        If .RecordCount = 0 Then
            GeraVenda = ""
            Exit Function
        End If
        
        num_ven = vendas.F_GeraNovaVenda(!cod_cli, !cod_For, !cnpj_Cli, "V", cod_emp)
        If num_ven = "" Then
            Exit Function
        End If
        
        vendas.AtualizaVenda num_ven, CStr(!valor_final), CStr(!desconto), !cliente, !telefone, !contato, !cod_cli, _
            strCod_Cai, strCod_Usu, TiraNull(!observacao), TiraNull(!email)
        
    
    End With
    
    
    
    
    strSQL = "select i.* from orcamento_itens  as i "
    strSQL = strSQL & " where 1>0 "
    strSQL = strSQL & " and i.id_orcamento=" & id_orcamento
    abreObjetoAdo rItens, strSQL
    Dim seq As String
    Dim p As Produto
    With rItens
        Do While Not .EOF
            p.cod_pro = !cod_pro
            p.nome_pro = !Produto
            p.pv_pro = !unitario
            p.detalhe.bordado = TiraNull(!bordado)
            p.detalhe.cor = TiraNull(!cor)
            p.detalhe.detalhamento = TiraNull(!detalhamento)
            p.detalhe.silk = TiraNull(!silk)
            
            
            
            seq = vendas.F_AdicionaItem(num_ven, p, !Quantidade, "0", CStr(!total), "0", "01", "N", !detalhamento)
            
            .MoveNext
        Loop
    End With
    
    
    rPesq.Close
    rItens.Close
    GeraVenda = num_ven
    
    Exit Function
trata_Erro:
    
End Function


Function DetalhamentoUniforme(cod_pro As String, Optional cor As String = "", Optional bordado As String = "", _
    Optional silk As String = "", Optional detalhamento As String = "") As detalhamento_uniforme
    frmDetalheProdutoUniformes.tag = cod_pro
    frmDetalheProdutoUniformes.txtBordado.Text = bordado
    frmDetalheProdutoUniformes.txtCor.Text = cor
    frmDetalheProdutoUniformes.txtSilk.Text = silk
    frmDetalheProdutoUniformes.txtDetalhamento.Text = detalhamento
    frmDetalheProdutoUniformes.Show 1
    If frmDetalheProdutoUniformes.tag <> "" Then
        DetalhamentoUniforme.bordado = frmDetalheProdutoUniformes.txtBordado.Text
        DetalhamentoUniforme.cor = frmDetalheProdutoUniformes.txtCor.Text
        DetalhamentoUniforme.detalhamento = frmDetalheProdutoUniformes.txtDetalhamento.Text
        DetalhamentoUniforme.silk = frmDetalheProdutoUniformes.txtSilk.Text
        Unload frmDetalheProdutoUniformes
    End If
    
    
    
End Function

