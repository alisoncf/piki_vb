Attribute VB_Name = "DW"
Public bDw As database
Function F_GravaMovProd(cod_pro As String, sai1 As String, sai2 As String, ent1 As String, ent2 As String, pv As String, cr As String)
    Dim strSQL As String
    strSQL = "insert into tab_movprod(cod,sai1,sai2,ent1,ent2,pv,cr)values"
    
End Function
Sub P_GravaMovMensal(mes As String, ano As String, avista As String, _
aprazo As String, financeira As String, impfed As String, _
impest As String, saida As String, saida2 As String, cmp1 As String, cmp2 As String)
    Dim strSQL As String
    strSQL = "INSERT INTO movmensal (mes, ano, avista, aprazo, financeira, impfed, impest, saida,saida2,cmp1,cmp2)values("
    strSQL = strSQL & mes & ","
    strSQL = strSQL & ano & ","
    strSQL = strSQL & F_GravaDec(avista) & ","
    strSQL = strSQL & F_GravaDec(aprazo) & ","
    strSQL = strSQL & F_GravaDec(financeira) & ","
    strSQL = strSQL & F_GravaDec(impfed) & ","
    strSQL = strSQL & F_GravaDec(impest) & ","
    strSQL = strSQL & F_GravaDec(saida) & ","
    strSQL = strSQL & F_GravaDec(saida2) & ","
    strSQL = strSQL & F_GravaDec(cmp1) & ","
    strSQL = strSQL & F_GravaDec(cmp2) & ")"
    conn.Execute strSQL
    
    
End Sub


Sub P_GravaProdutoPreco(codigo As String, nome As String, customax As String, customin As String, Entrada As String, saida As String, _
    entrada2 As String, saida2 As String, er As String, ef As String, mes As String, ano As String, preco As String)
    Dim strSQL As String
    strSQL = "INSERT INTO produtopreco (codigo,nome, customax, customin, entrada, saida, "
    strSQL = strSQL & "entrada2, saida2, er, ef, preco, mes, ano)values("
    strSQL = strSQL & "'" & codigo & "',"
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & F_GravaDec(customax) & ","
    strSQL = strSQL & F_GravaDec(customin) & ","
    strSQL = strSQL & F_GravaDec(Entrada) & ","
    strSQL = strSQL & F_GravaDec(saida) & ","
    strSQL = strSQL & F_GravaDec(entrada2) & ","
    strSQL = strSQL & F_GravaDec(saida2) & ","
    strSQL = strSQL & F_GravaDec(er) & ","
    strSQL = strSQL & F_GravaDec(ef) & ","
    strSQL = strSQL & F_GravaDec(preco) & ","
    strSQL = strSQL & mes & ","
    strSQL = strSQL & ano & ")"
    banco.Execute strSQL
    
    
    
    
    
End Sub
Sub P_GravaProdutoVenda(codigo As String, nome As String, mes As String, ano As String, qtd As String, valor As String, saida As String)
    Dim strSQL As String
    strSQL = "INSERT INTO produtovenda (codigo, nome, mes, ano, qtd, valor, saida) VALUES("
    strSQL = strSQL & "'" & codigo & "',"
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & mes & ","
    strSQL = strSQL & ano & ","
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & F_GravaDec(saida) & ")"
    banco.Execute strSQL
    
    
    
End Sub
