Attribute VB_Name = "Sped_Util"
Function Sped_Lista_Cst_Pis(expr As String) As Padrao1
    Sped_Lista_Cst_Pis = ListaPadrao1(expr, "sped_cst_pis", "codigo", "descricao", "tipo")
    
End Function
Function Sped_Lista_Cst_Icms(expr As String) As Padrao1
    Sped_Lista_Cst_Icms = ListaPadrao1(expr, "nfe_tipo_icms", "codigo", "descricao", "tabela")
    
End Function
Function Sped_Lista_Cst_Cofins(expr As String) As Padrao1
    Sped_Lista_Cst_Cofins = ListaPadrao1(expr, "sped_cst_cofins", "codigo", "descricao")
    
End Function

Function Sped_Lista_Cst_IPI(expr As String) As Padrao1
    Sped_Lista_Cst_IPI = ListaPadrao1(expr, "sped_cst_ipi", "codigo", "descricao")
    
End Function
Function Sped_Tipo_Item(expr As String) As Padrao1
    Sped_Tipo_Item = ListaPadrao1(expr, "sped_tipo_item", "codigo", "descricao")
    
End Function
Function Sped_Situacao_Doc(expr As String) As Padrao1
    Sped_Situacao_Doc = ListaPadrao1(expr, "sped_situacao_doc", "codigo", "descricao")
    
End Function
Function Sped_Genero(expr As String) As Padrao1
    Sped_Genero = ListaPadrao1(expr, "sped_genero", "codigo", "descricao")
    
End Function
Function Sped_Lista_CFOP(expr As String) As Padrao1
    Sped_Lista_CFOP = ListaPadrao1(expr, "sped_CFOP", "codigo", "descricao")
    
End Function
Function Sped_Origem(expr As String) As Padrao1
    Sped_Origem = ListaPadrao1(expr, "sped_origem", "codigo", "descricao", "tipo")
    
End Function
