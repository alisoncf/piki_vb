Attribute VB_Name = "Cidades"
Public Type Cidade
    nome_cid As String
    cod_cid As String
    uf_cid As String
    cod_mun As String
End Type


Public Function F_GravaCidade(nome_cid As String, uf_cid As String) As String
    Dim sCod As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    sCod = F_BuscaUmCampo("cod_cid", TAB_CIDADES, "nome_cid='" & nome_cid & "'")
    If sCod = "" Then
        sCod = MaxKeySQL("tab_cidades", "cod_cid", "", "0000#")
    
    
        strSQL = "insert into " & TAB_CIDADES & " (cod_cid,nome_cid,uf_cid)values("
        strSQL = strSQL & "'" & sCod & "',"
        strSQL = strSQL & "'" & nome_cid & "',"
        strSQL = strSQL & "'" & uf_cid & "')"
        Banco.Execute strSQL
    End If
    F_GravaCidade = sCod
    Exit Function
trata_Erro:
    
End Function

Public Function F_AlteraCidade(cod_cid As String, nome_cid As String, uf_cid As String) As Boolean
    Dim sCod As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = " update " & TAB_CIDADES & " set"
    
    strSQL = strSQL & " nome_cid = '" & nome_cid & "',"
    strSQL = strSQL & "uf_cid='" & uf_cid & "' "
    strSQL = strSQL & " where cod_cid='" & cod_cid & "' "
    Banco.Execute strSQL
    F_AlteraCidade = True
    Exit Function
trata_Erro:
    
End Function
Function F_BuscaCidades(expr As String) As Cidade
    On Error GoTo trata_Erro
    frmEdiCid.tag = "B"
    frmEdiCid.txtbxNOme.text = expr
    frmEdiCid.Show 1
    
    If frmEdiCid.tag <> "" Then
        F_BuscaCidades.cod_cid = frmEdiCid.adc.Recordset!cod_cid
        F_BuscaCidades.nome_cid = frmEdiCid.adc.Recordset!nome_cid
        F_BuscaCidades.uf_cid = frmEdiCid.adc.Recordset!uf_cid
    End If

trata_Erro:

End Function
Function F_BuscaCidadeCodigo(codigo As String) As Cidade
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from " & TAB_CIDADES & " where cod_cid='" & codigo & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            F_BuscaCidadeCodigo.cod_cid = !cod_cid
            F_BuscaCidadeCodigo.nome_cid = !nome_cid
            F_BuscaCidadeCodigo.uf_cid = !uf_cid
            .MoveNext
        Loop
        .Close
    End With
    
End Function
Sub P_PreencheCidades(obj As Object, Optional UF As String = "")
    Dim sWhere As String
    If UF <> "" Then
        sWhere = "uf_cid='" & UF & "' "
    End If
    P_Preenche obj, TAB_CIDADES, "nome_cid", sWhere
    

End Sub

