Attribute VB_Name = "Fornecedor"

Type fornecedor
    cod_For As String
    nome_for As String
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
    ie_for As String
End Type

Function F_SelectForn(expr As String) As fornecedor.fornecedor
    Dim strSQL As String
    strSQL = "select cod_for as cod,nome_for as nome,fantasia_for as fantasia,ie_for as ie,"
    strSQL = strSQL & " cnpj_for as cnpj,"
    strSQL = strSQL & "endereco_for as endereco,"
    strSQL = strSQL & "cidade_for as Cidade,"
    strSQL = strSQL & "uf_for as uf,"
    strSQL = strSQL & "telefone_for as Telefone,"
    strSQL = strSQL & "fax_for as fax,"
    strSQL = strSQL & "email_for as email,"
    strSQL = strSQL & "contato_for as contato "
    strSQL = strSQL & " from tab_fornecedores "
    
    frmSelect.cCriterio.Add " where (nome_for like '" & expr & "%' or Fantasia_for like '%" & expr & "%')"
    
    
    frmSelect.txtbxExpr.Text = expr
    frmSelect.sOrd = " order by fantasia_for "
    
    
    
    
    frmSelect.sSelect = strSQL
    
    frmSelect.Show 1
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        F_SelectForn.cod_For = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 1
        F_SelectForn.nome_for = frmSelect.msflxgrd
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

Sub F_GravaFornCurto(cod_For As String, nome_for As String, fantasia_for As String)
   Dim strSQL As String
   On Error GoTo trata_Erro
   strSQL = "insert into tab_fornecedores(cod_for,nome_for,fantasia_for)values("
   strSQL = strSQL & "'" & cod_For & "',"
   strSQL = strSQL & "'" & nome_for & "',"
   strSQL = strSQL & "'" & fantasia_for & "')"
   Conn.Execute strSQL
   Exit Sub
trata_Erro:
   P_MostraErro
End Sub
Function F_GravaFor(nome As String, fantasia As String, CNPJ As String, IE As String, Endereco As String, _
    bairro As String, Cidade As String, UF As String, cep As String, contato As String, _
    email As String, obs As String)
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    Dim iCount As Integer
    Dim sTel As String
    Dim sTipo As String
    
    sCod = MaxKeySQL("tab_fornecedores", "cod_for", "", "0000#")
    
    
    
    
    strSQL = "insert into tab_fornecedores(cod_for,nome_for,fantasia_for,cnpj_for,ie_for,endereco_for, "
    strSQL = strSQL & "bairro_for,cidade_for,uf_for,cep_for,contato_for,email_for,obs_for) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & "'" & fantasia & "',"
    strSQL = strSQL & "'" & CNPJ & "',"
    strSQL = strSQL & "'" & IE & "',"
    strSQL = strSQL & "'" & Endereco & "',"
    strSQL = strSQL & "'" & bairro & "',"
    strSQL = strSQL & "'" & Cidade & "',"
    strSQL = strSQL & "'" & UF & "',"
    strSQL = strSQL & "'" & cep & "',"
    'strSQL = strSQL & "'" & txtbxTel.Text & "',"
    'strSQL = strSQL & "'" & txtbxFax.Text & "',"
    strSQL = strSQL & "'" & contato & "',"
    strSQL = strSQL & "'" & email & "',"
    strSQL = strSQL & "'" & obs & "') "
    
    Conn.Execute strSQL
    
    F_GravaFor = sCod
    
    
Exit Function
trata_Erro:
    MsgBox Err.Description, vbCritical
End Function

Function F_BuscaFornecedor(expr As String) As fornecedor.fornecedor
    frmEdiFor.tag = "x"
    frmEdiFor.WindowState = 2
    frmEdiFor.txtbxExpr.Text = expr
    frmEdiFor.Show 1
    
        
    If frmEdiFor.tag <> "" Then
        With frmEdiFor.ado
            F_BuscaFornecedor.bairro_for = TiraNull(!bairro_for)
            F_BuscaFornecedor.cidade_For = TiraNull(!cidade_For)
            F_BuscaFornecedor.cnpj_for = TiraNull(!cnpj_for)
            F_BuscaFornecedor.cod_For = !cod_For
            F_BuscaFornecedor.contato_for = TiraNull(!contato_for)
            F_BuscaFornecedor.email_for = TiraNull(!email_for)
            F_BuscaFornecedor.endereco_for = TiraNull(!endereco_for)
            F_BuscaFornecedor.fantasia_for = TiraNull(!fantasia_for)
            F_BuscaFornecedor.fax_for = TiraNull(!fax_for)
            F_BuscaFornecedor.nome_for = TiraNull(!nome_for)
            F_BuscaFornecedor.telefone_for = TiraNull(!telefone_for)
            F_BuscaFornecedor.uf_for = TiraNull(!uf_for)
        End With
    End If
    Unload frmEdiFor
    
End Function

