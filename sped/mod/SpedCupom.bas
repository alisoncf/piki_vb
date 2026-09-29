Attribute VB_Name = "SpedCupom"
Dim strSQL As String
    Public Type sped_cupom
        id As String
        cod_mod As String
        cod_sit As String
        num_doc As String
        vl_doc As String
        vl_pis As String
        vl_cofins As String
        cpf_cnpq As String
        nom_adq As String
    End Type

Public Type sped_redz
    id As String
    caixa As String
    dt_doc As String
    cro As String
    crz As String
    num_coo_fin As String
    gt_fin As String
    vl_brt As String

End Type

Public Function gravasped_cupom(num_ven As String, caixa As String, data As String, hora As String, cod_mod As String, cod_sit As String, num_doc As String, vl_doc As String, vl_pis As String, vl_cofins As String, cpf_apq As String, nom_adq As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    
    If cod_mod = "" Then
        cod_mod = "2D"
    
    End If
    If cod_sit = "" Then
        cod_sit = "00" 'documento regular - vide tabela sped_situacao_doc
    End If
    
    strSQL = " insert into sped_cupom(caixa,num_ven,data,hora,cod_mod,cod_sit,num_doc,vl_doc,vl_pis,vl_cofins,cpf_adq,nom_adq)values("
    strSQL = strSQL & "'" & caixa & "',"
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & F_GravaData(data) & ","
    strSQL = strSQL & "'" & hora & "',"
    strSQL = strSQL & "'" & cod_mod & "',"
    strSQL = strSQL & "'" & cod_sit & "',"
    strSQL = strSQL & "'" & num_doc & "',"
    strSQL = strSQL & str(vl_doc) & ","
    strSQL = strSQL & str(vl_pis) & ","
    strSQL = strSQL & str(vl_cofins) & ","
    strSQL = strSQL & "'" & cpf_adq & "',"
    strSQL = strSQL & "'" & nom_adq & "'"
    strSQL = strSQL & ")"
    conn.Execute strSQL
    strSQL = "set @id_cupom=(select max(id) from sped_cupom)"
    conn.Execute strSQL
    gravasped_cupom = F_BuscaUmCampo("@id_cupom", "", "")
    
    rPesq.Open "select @id_cupom as id_cupom", conn
    gravasped_cupom = rPesq!id_cupom
    rPesq.Close
    Exit Function
trata_Erro:
    

End Function

Sub Alterasped_cupom(id As String, caixa As String, data As String, hora As String, cod_mod As String, cod_sit As String, num_doc As String, vl_doc As String, vl_pis As String, vl_cofins As String, cpf_cnpq As String, nom_adq As String)
    strSQL = " update sped_cupom set cod_mod='" & cod_mod & "',"
    strSQL = strSQL & "cod_sit='" & cod_sit & "',"
    strSQL = strSQL & "caixa='" & caixa & "',"
    strSQL = strSQL & "data='" & F_GravaData(Date) & ","
    strSQL = strSQL & "hora='" & hora & "',"
    strSQL = strSQL & "num_doc='" & num_doc & "',"
    strSQL = strSQL & "vl_doc=" & str(vl_doc) & ","
    strSQL = strSQL & "vl_pis=" & str(vl_pis) & ","
    strSQL = strSQL & "vl_cofins=" & str(vl_cofins) & ","
    strSQL = strSQL & "cpf_cnpq='" & cpf_cnpq & "',"
    strSQL = strSQL & "nom_adq='" & nom_adq & "'"
    strSQL = strSQL & " where id=" & id
    conn.Execute strSQL
End Sub

Public Sub gravasped_cupom_itens(id_cupom As String, cod_item As String, qtd As String, _
    qtd_canc As String, unid As String, vl_item As String, cst_icms As String, CFOP As String, _
    aliq_icms As String, vl_pis As String, vl_cofins As String, vl_icms As String, vl_desconto As String)
    strSQL = " insert into sped_cupom_itens(id_cupom,cod_item,qtd,qtd_canc,unid,vl_item,cst_icms,cfop,"
    strSQL = strSQL & "aliq_icms,vl_pis,vl_cofins,vl_icms,vl_desconto)values("
    strSQL = strSQL & "'" & id_cupom & "',"
    strSQL = strSQL & "'" & cod_item & "',"
    strSQL = strSQL & str(qtd) & ","
    strSQL = strSQL & str(qtd_canc) & ","
    strSQL = strSQL & "'" & unid & "',"
    strSQL = strSQL & str(vl_item) & ","
    strSQL = strSQL & "'" & cst_icms & "',"
    strSQL = strSQL & "'" & CFOP & "',"
    strSQL = strSQL & str(aliq_icms) & ","
    strSQL = strSQL & str(vl_pis) & ","
    strSQL = strSQL & str(vl_cofins) & ","
    strSQL = strSQL & str(vl_icms) & ","
    strSQL = strSQL & str(vl_desconto) & ")"
    conn.Execute strSQL

End Sub

Sub Alterasped_cupom_itens(id As String, id_cupom As String, cod_item As String, qtd As String, qtd_canc As String, _
unid As String, vl_item As String, cst_icms As String, CFOP As String, aliq_icms As String, vl_pis As String, _
vl_cofins As String, vl_icms As String, vl_desconto As String)
    strSQL = " update sped_cupom_itens set "
    strSQL = strSQL & "id_cupom=" & id_cupom & ","
    strSQL = strSQL & "cod_item='" & cod_item & "',"
    strSQL = strSQL & "qtd=" & str(qtd) & ","
    strSQL = strSQL & "qtd_canc=" & str(qtd_canc) & ","
    strSQL = strSQL & "unid='" & unid & "',"
    strSQL = strSQL & "vl_item=" & str(vl_item) & ","
    strSQL = strSQL & "cst_icms='" & cst_icms & "',"
    strSQL = strSQL & "cfop='" & CFOP & "',"
    strSQL = strSQL & "aliq_icms=" & str(aliq_icms) & ","
    strSQL = strSQL & "vl_pis=" & str(vl_pis) & ","
    strSQL = strSQL & "vl_icms=" & str(vl_icms) & ","
    strSQL = strSQL & "vl_desconto=" & str(vl_desconto) & ","
    strSQL = strSQL & "vl_cofins=" & str(vl_cofins) & " "
    strSQL = strSQL & " where id=" & id
    conn.Execute strSQL
End Sub

        


Function GravaSpedRedZ(caixa As String, dt_doc As String, cro As String, crz As String, num_coo_fin As String, _
    gt_fin As String, vl_brt As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro:
    
    strSQL = "Insert into sped_redz(caixa, dt_doc,cro,crz,num_coo_fin,gt_fin,vl_brt)values("
    strSQL = strSQL & "'" & caixa & "',"
    strSQL = strSQL & F_GravaData(dt_doc) & ","
    strSQL = strSQL & F_GravaDec(cro) & ","
    strSQL = strSQL & F_GravaDec(crz) & ","
    strSQL = strSQL & F_GravaDec(num_coo_fin) & ","
    strSQL = strSQL & F_GravaDec(gt_fin) & ","
    strSQL = strSQL & F_GravaDec(vl_brt) & ")"
    conn.Execute strSQL
    
    GravaSpedRedZ = True
    Exit Function
trata_Erro:
    GravaSpedRedZ = False
End Function
Function AlteraSpedRedZTot(id As String, id_red As String, valor As String, cod_imp As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro:
    
    strSQL = "update sped_redz_tot set "
    strSQL = strSQL & "id_red=" & id_red & ","
    
    strSQL = strSQL & "valor=" & F_GravaDec(valor) & ","
    strSQL = strSQL & "cod_imp='" & cod_imp & "' "
    strSQL = strSQL & " where id=" & id
    conn.Execute strSQL
    
    AlteraSpedRedZTot = True
    Exit Function
trata_Erro:
    AlteraSpedRedZTot = False
End Function
Function SelectSpedRedZTot(id_red As String) As String
    Dim strSQL As String
    strSQL = "select * from sped_redz_tot(id_red, valor,cod_imp)values("
    strSQL = strSQL & id_red & ","
    
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & "'" & cod_imp & "') "
    conn.Execute strSQL
    
End Function
Function GravaSpedRedZTot(id_red As String, valor As String, cod_imp As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro:
    
    strSQL = "Insert into sped_redz_tot(id_red, valor,cod_imp)values("
    strSQL = strSQL & id_red & ","
    
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & "'" & cod_imp & "') "
    conn.Execute strSQL
    
    GravaSpedRedZTot = True
    Exit Function
trata_Erro:
    GravaSpedRedZTot = False
End Function

Function AlteraSpedRedZ(id As String, caixa As String, dt_doc As String, cro As String, crz As String, num_coo_fin As String, _
    gt_fin As String, vl_brt As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro:
    
    strSQL = "update sped_redz set "
    strSQL = strSQL & "caixa='" & caixa & "',"
    strSQL = strSQL & "dt_doc=" & F_GravaData(dt_doc) & ","
    strSQL = strSQL & "cro=" & F_GravaDec(cro) & ","
    strSQL = strSQL & "crz=" & F_GravaDec(crz) & ","
    strSQL = strSQL & "num_coo_fin=" & F_GravaDec(num_coo_fin) & ","
    strSQL = strSQL & "gt_fin=" & F_GravaDec(gt_fin) & ","
    strSQL = strSQL & "vl_brt=" & F_GravaDec(vl_brt) & ""
    strSQL = strSQL & " where id=" & id
    conn.Execute strSQL
    
    AlteraSpedRedZ = True
    Exit Function
trata_Erro:
    AlteraSpedRedZ = False
End Function

Function BuscaSpedRedZ(dataI As String, Optional dataF As String, Optional caixa As String) As String
    Dim strSQL As String
    strSQL = "select * from sped_redz where dt_doc >=" & F_GravaData(dataI)
    If dataF <> "" Then
        strSQL = strSQL & " and dt_doc<=" & F_GravaData(dataF)
    End If
    If caixa <> "" Then
        strSQL = strSQL & " and caixa='" & caixa & "'"
    End If
    
    BuscaSpedRedZ = strSQL
End Function
Sub MigraCupom(dataI As String, dataF As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset

End Sub
