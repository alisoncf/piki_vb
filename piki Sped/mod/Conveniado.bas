Attribute VB_Name = "Conveniado"
Public Type conveniado
    cod_cli As String
    seq_con As String
    nome_con As String
    endereco_con As String
    bairro_con As String
    cidade_con As String
    uf_con As String
    fone_con As String
    cpf_con As String


End Type


Function F_GravaConveniado(cod_cli As String, nome_con As String, endereco_con As String, _
    bairro_con As String, uf_con As String, fone_con As String, cpf_con As String, matricula_con) As String
    Dim sSeq As String
    Dim strSQL As String
    sSeq = MaxKeySQL("tab_conveniados", "seq_con", "", "")
    strSQL = "insert into tab_Conveniados(cod_cli,seq_con,nome_con,endereco_con,bairro_con,uf_con,fone_con,cpf_con,matricula_con)values("
    strSQL = strSQL & "'" & cod_cli & "', "
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & "'" & nome_con & "',"
    strSQL = strSQL & "'" & endereco_con & "',"
    strSQL = strSQL & "'" & bairro_con & "',"
    strSQL = strSQL & "'" & uf_con & "',"
    strSQL = strSQL & "'" & fone_con & "',"
    strSQL = strSQL & "'" & cpf_con & "', "
    strSQL = strSQL & "'" & matricula_con & "') "
    Conn.Execute strSQL
    
    F_GravaConveniado = sSeq
    
End Function
    

Function F_AlteraConveniado(seq_con, nome_con As String, endereco_con As String, _
    bairro_con As String, uf_con As String, fone_con As String, cpf_con As String) As String
    Dim sSeq As String
    Dim strSQL As String
    strSQL = "update tab_Conveniados set "
    
    strSQL = strSQL & "nome_CON='" & nome_con & "',"
    strSQL = strSQL & "endereco_con='" & endereco_con & "',"
    strSQL = strSQL & "bairro_con='" & bairro_con & "',"
    strSQL = strSQL & "uf_con='" & uf_con & "',"
    strSQL = strSQL & "fone_con='" & fone_con & "',"
    strSQL = strSQL & "celular_con='" & cpf_con & "' "
    
    strSQL = strSQL & " where seq_con='" & cod_con & "' "
    
    
    Conn.Execute strSQL
    
    
    
End Function
    
Function F_SelectConveniado(sExpr As String, cod_cli As String) As conveniado.conveniado
    
    frmSelect.sSelect = "select seq_con as con,nome_con as nome,cpf_con as cpf, " & _
        "endereco_con as endereco,bairro_con as bairro, cidade_con as cidade,uf_con as uf, fone_con as fone,cod_cli as cod_cli " & _
        " from tab_conveniados where cod_cli='" & cod_cli & "' "
    
    frmSelect.cCriterio.Add " and nome_con like '%" & sExpr & "%'"
    frmSelect.sOrd = " order by nome_con "
    frmSelect.txtbxExpr.Text = expr
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.Col = 0
        F_SelectConveniado.seq_con = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.Col = 1
        F_SelectConveniado.nome_con = frmSelect.msflxgrd
        
        frmSelect.msflxgrd.Col = 2
        F_SelectConveniado.cpf_con = frmSelect.msflxgrd
        
        frmSelect.msflxgrd.Col = 3
        F_SelectConveniado.endereco_con = frmSelect.msflxgrd
        frmSelect.msflxgrd.Col = 4
        F_SelectConveniado.bairro_con = frmSelect.msflxgrd
        frmSelect.msflxgrd.Col = 5
        F_SelectConveniado.cidade_con = frmSelect.msflxgrd
        frmSelect.msflxgrd.Col = 6
        F_SelectConveniado.uf_con = frmSelect.msflxgrd
        frmSelect.msflxgrd.Col = 7
        F_SelectConveniado.fone_con = frmSelect.msflxgrd
        frmSelect.msflxgrd.Col = 8
        F_SelectConveniado.cod_cli = frmSelect.msflxgrd
        
        
    End If
    Unload frmSelect
    

End Function


Function F_BuscaConveniado(seq_con As String, cod_cli As String) As conveniado.conveniado
    
End Function
Function f_BuscaConveniadoNome(seq_con As String, cod_cli As String) As String
    f_BuscaConveniadoNome = F_BuscaUmCampo("nome_con", "tab_conveniados", "cod_cli='" & cod_cli & "' " & _
        " AND seq_con=" & seq_con)
End Function
