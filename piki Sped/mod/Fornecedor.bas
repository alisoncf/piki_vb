Attribute VB_Name = "Fornecedor"


Type fornecedor
    cod_for As String
    nome_For As String
    fantasia_for As String
    cnpj_for As String
    endereco_for As String
    bairro_for As String
    cidade_For As String
    uf_for As String
    telefone_for As String
    fax_for As String
    email_for As String
    contato_for As String
    ie As String
End Type

Function F_SelectForn(expr As String) As fornecedor.fornecedor
    Dim strSQL As String
    strSQL = "select cod_for as cod,nome_for as nome,fantasia_for as fantasia,"
    strSQL = strSQL & " cnpj_for as cnpj,"
    strSQL = strSQL & "endereco_for as endereco,"
    strSQL = strSQL & "cidade_for as Cidade,"
    strSQL = strSQL & "uf_for as uf,"
    strSQL = strSQL & "telefone_for as Telefone,"
    strSQL = strSQL & "fax_for as fax,"
    strSQL = strSQL & "email_for as email,"
    strSQL = strSQL & "contato_for as contato "
    strSQL = strSQL & " from " & TAB_FORNECEDORES & " "
    
    frmSelect.cCriterio.Add " where (nome_for like '" & expr & "%' or Fantasia_for like '%" & expr & "%') and visivel_for<>'N'"
    
    
    frmSelect.txtbxExpr.Text = expr
    frmSelect.sOrd = " order by fantasia_for "
    
    
    
    
    frmSelect.sSelect = strSQL
    
    frmSelect.Show 1
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        F_SelectForn.cod_for = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 1
        F_SelectForn.nome_For = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 2
        F_SelectForn.fantasia_for = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 3
        F_SelectForn.cnpj_for = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 4
        F_SelectForn.endereco_for = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 5
        F_SelectForn.cidade_For = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 6
        F_SelectForn.uf_for = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 7
        F_SelectForn.telefone_for = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 8
        F_SelectForn.fax_for = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 9
        F_SelectForn.email_for = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 10
        F_SelectForn.contato_for = frmSelect.msflxgrd
    End If
    Unload frmSelect
    
End Function
Function F_BuscaFornecedor(expr As String) As fornecedor.fornecedor
        frmEdiFor.tag = "x"
    frmEdiFor.WindowState = 2
    frmEdiFor.txtbxExpr.Text = expr
    frmEdiFor.Show 1
    
        
    If frmEdiFor.tag <> "" Then
        With frmEdiFor.adc.Recordset
            F_BuscaFornecedor.bairro_for = !bairro_for
            F_BuscaFornecedor.cidade_For = !cidade_For
            F_BuscaFornecedor.cnpj_for = !cnpj_for
            F_BuscaFornecedor.cod_for = !cod_for
            F_BuscaFornecedor.contato_for = !contato_for
            F_BuscaFornecedor.email_for = !email_for
            F_BuscaFornecedor.endereco_for = !endereco_for
            F_BuscaFornecedor.fantasia_for = !fantasia_for
            F_BuscaFornecedor.fax_for = !fax_for
            F_BuscaFornecedor.nome_For = !nome_For
            F_BuscaFornecedor.telefone_for = !telefone_for
            F_BuscaFornecedor.uf_for = !uf_for
        End With
    End If
    Unload frmEdiFor
    
End Function
Sub F_GravaFornCurto(cod_for As String, nome_For As String, fantasia_for As String)
   Dim strSQL As String
   On Error GoTo trata_Erro
   strSQL = "insert into " & TAB_FORNECEDORES & "(cod_for,nome_for,fantasia_for)values("
   strSQL = strSQL & "'" & cod_for & "',"
   strSQL = strSQL & "'" & nome_For & "',"
   strSQL = strSQL & "'" & fantasia_for & "')"
   conn.Execute strSQL
   Exit Sub
trata_Erro:
   MostraErro
End Sub
Function F_GravaFor(nome As String, fantasia As String, CNPJ As String, ie As String, Endereco As String, _
    bairro As String, Cidade As String, UF As String, cep As String, contato As String, _
    email As String, telefone As String, fax As String, obs As String, Optional cod_for As String = "")
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    Dim iCount As Integer
    Dim sTel As String
    Dim sTipo As String
    
    If cod_for = "" Then
        sCod = MaxKeySQL(TAB_FORNECEDORES, "cod_for", "", "0000#")
    Else
        sCod = cod_for
    End If
    
    
    
    strSQL = "insert into " & TAB_FORNECEDORES & "(cod_for,nome_for,fantasia_for,cnpj_for,ie_for,endereco_for, "
    strSQL = strSQL & "bairro_for,cidade_for,uf_for,cep_for,telefone_for,fax_For,contato_for,email_for,visivel_For,obs_for) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & "'" & fantasia & "',"
    strSQL = strSQL & "'" & CNPJ & "',"
    strSQL = strSQL & "'" & ie & "',"
    strSQL = strSQL & "'" & Endereco & "',"
    strSQL = strSQL & "'" & bairro & "',"
    strSQL = strSQL & "'" & Cidade & "',"
    strSQL = strSQL & "'" & UF & "',"
    strSQL = strSQL & "'" & cep & "',"
    strSQL = strSQL & "'" & telefone & "',"
    strSQL = strSQL & "'" & fax & "',"
    strSQL = strSQL & "'" & contato & "',"
    strSQL = strSQL & "'" & email & "',"
    strSQL = strSQL & "'S',"
    strSQL = strSQL & "'" & obs & "') "
    
    conn.Execute strSQL
    
    F_GravaFor = sCod
    
    
Exit Function
trata_Erro:
    MsgBox Err.Description, vbCritical
End Function




