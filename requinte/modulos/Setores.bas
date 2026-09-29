Attribute VB_Name = "Setores"
Type setor
    cod As String
    nome As String
    Cidade As String
    cobrador_set As String
    UF As String
    tipo As String
End Type
Function F_gravaSetor(tipo As String, nome As String, Cidade As String, cod_fun As String) As String
    Dim sCod As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    sCod = MaxKeySQL("tab_setores", "cod_set", "", "0000#")
    strSQL = "insert into tab_setores(cod_set,tipo_set,nome_set,cod_cid,cobrador_set)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & tipo & "',"
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & "'" & Cidade & "',"
    strSQL = strSQL & "'" & cod_fun & "')"
    Conn.Execute strSQL
    F_gravaSetor = sCod
    Exit Function
trata_Erro:
    
End Function

Function F_EditaSetor(cod As String, tipo As String, nome As String, Cidade As String, cod_fun As String) As Boolean
    Dim sCod As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "update tab_setores  set "
    strSQL = strSQL & "tipo_set='" & tipo & "',"
    strSQL = strSQL & "nome_set='" & nome & "',"
    strSQL = strSQL & "cod_cid='" & Cidade & "',"
    strSQL = strSQL & "cobrador_set='" & cod_fun & "' "
    strSQL = strSQL & " WHERE cod_set='" & cod & "' "
    Conn.Execute strSQL
    F_EditaSetor = True
    Exit Function
trata_Erro:
    
End Function

Function F_BuscaSetor(expr As String, Optional cod_cid As String = "") As setor
    On Error GoTo trata_Erro
    frmEdiSet.tag = "B"
    frmEdiSet.caption = "Escolha um setor"
    frmEdiSet.txtbxBairro.text = expr
    frmEdiSet.txtbxCidade.tag = cod_cid
    frmEdiSet.Show 1
    
    If frmEdiSet.tag <> "" Then
        F_BuscaSetor.Cidade = frmEdiSet.ado!nome_cid
        F_BuscaSetor.cobrador_set = frmEdiSet.ado!cobrador_set
        F_BuscaSetor.cod = frmEdiSet.ado!cod_set
        F_BuscaSetor.nome = frmEdiSet.ado!nome_set
        F_BuscaSetor.UF = frmEdiSet.ado!uf_cid
        F_BuscaSetor.tipo = frmEdiSet.ado!tipo_set
    End If
    
    
Exit Function
trata_Erro:
End Function
Function F_BuscaSetorPorCodigo(cod_set As String) As setor
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select ts.*,tc.nome_cid,tc.uf_cid from tab_cidades as tc,"
    strSQL = strSQL & " tab_setores as ts "
    strSQL = strSQL & " where tc.cod_cid=ts.cod_cid "
    
    strSQL = strSQL & " and cod_set='" & cod_set & "'     "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            F_BuscaSetorPorCodigo.Cidade = !nome_cid
            F_BuscaSetorPorCodigo.cobrador_set = !cobrador_set
            F_BuscaSetorPorCodigo.cod = !cod_set
            F_BuscaSetorPorCodigo.nome = !nome_set
            F_BuscaSetorPorCodigo.UF = !uf_cid
            F_BuscaSetorPorCodigo.tipo = !tipo_set
        End If
        .Close
    End With
    
    
Exit Function
trata_Erro:
End Function



Function F_RetornaSetorTrab(cod_cli As String) As setor
    Dim sCod As String
    
    sCod = F_BuscaUmCampo("cod_set", "tab_Clientetrab", "cod_cli='" & cod_cli & "'")
    If sCod <> "" Then
        F_RetornaSetorTrab = F_RetornaSetor(sCod)
    End If
    
    
    
End Function
Function F_RetornaSetor(codigo As String) As setor
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select ts.*,tc.nome_cid,tc.uf_cid from tab_setores as ts,tab_cidades as tc "
    strSQL = strSQL & " where ts.cod_cid=tc.cod_cid "
    strSQL = strSQL & " and ts.cod_set='" & codigo & "'"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        F_RetornaSetor.Cidade = !nome_cid
        F_RetornaSetor.cobrador_set = !cobrador_set
        F_RetornaSetor.cod = !cod_set
        F_RetornaSetor.nome = !nome_set
        F_RetornaSetor.tipo = !tipo_set
        F_RetornaSetor.UF = !uf_cid
        
        
        .Close
    End With
trata_Erro:
    
End Function
